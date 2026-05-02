CREATE EXTENSION IF NOT EXISTS pgcrypto;

CREATE OR REPLACE FUNCTION set_updated_at()
RETURNS TRIGGER AS $$
BEGIN
  NEW.updated_at = now();
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TYPE client_type AS ENUM ('person', 'company');
CREATE TYPE stock_movement_type AS ENUM (
  'receipt',
  'reserve',
  'unreserve',
  'issue_to_service',
  'install',
  'write_off',
  'sale',
  'return_from_customer',
  'return_to_stock',
  'transfer',
  'inventory_adjustment',
  'defect_write_off'
);
CREATE TYPE service_order_status AS ENUM (
  'accepted',
  'diagnostics',
  'awaiting_approval',
  'approval',
  'waiting_parts',
  'in_progress',
  'testing',
  'ready',
  'awaiting_payment',
  'issued',
  'closed',
  'cancelled',
  'cannot_repair',
  'waiting_client',
  'paused',
  'warranty_return'
);
CREATE TYPE service_part_status AS ENUM (
  'required',
  'to_purchase',
  'ordered',
  'in_transit',
  'arrived',
  'reserved',
  'issued_to_engineer',
  'installed',
  'written_off',
  'returned'
);
CREATE TYPE repair_payment_status AS ENUM (
  'unpaid',
  'prepayment',
  'partially_paid',
  'fully_paid',
  'refund',
  'debt'
);
CREATE TYPE purchase_order_status AS ENUM (
  'draft',
  'ordered',
  'in_transit',
  'partially_received',
  'received',
  'cancelled'
);
CREATE TYPE goods_receipt_status AS ENUM ('draft', 'posted', 'cancelled');
CREATE TYPE sale_status AS ENUM (
  'draft',
  'reserved',
  'awaiting_payment',
  'partially_paid',
  'paid',
  'issued',
  'closed',
  'cancelled',
  'returned'
);
CREATE TYPE payment_source_type AS ENUM ('sale', 'service_order');
CREATE TYPE payment_method AS ENUM ('cash', 'card_terminal', 'bank_transfer', 'mixed');
CREATE TYPE payment_kind AS ENUM ('prepayment', 'partial', 'full', 'additional_payment', 'refund');
CREATE TYPE cash_session_status AS ENUM ('open', 'closed');
CREATE TYPE cash_movement_type AS ENUM ('cash_in', 'cash_out', 'refund', 'correction');
CREATE TYPE inventory_session_status AS ENUM ('draft', 'in_progress', 'finished', 'cancelled');
CREATE TYPE price_type AS ENUM ('purchase', 'retail', 'wholesale', 'service_internal');
CREATE TYPE entity_type AS ENUM ('service_order', 'sale', 'client', 'product', 'purchase_order', 'goods_receipt');
CREATE TYPE document_kind AS ENUM (
  'intake_receipt',
  'diagnostic_act',
  'work_order',
  'completion_act',
  'device_issue_act',
  'sale_receipt',
  'sale_invoice',
  'cash_receipt',
  'return_document',
  'goods_receipt_invoice'
);
CREATE TYPE document_status AS ENUM ('draft', 'pdf_saved', 'signed', 'cancelled');
CREATE TYPE document_entity_type AS ENUM ('service_order', 'sale', 'purchase_order', 'goods_receipt', 'sale_return', 'payment');
CREATE TYPE audit_action_type AS ENUM (
  'create',
  'update',
  'delete',
  'status_change',
  'payment_accept',
  'issue_goods',
  'reserve_goods',
  'write_off',
  'login',
  'logout'
);
CREATE TYPE notification_type AS ENUM ('service', 'sale', 'purchase', 'stock', 'finance', 'system');
CREATE TYPE client_notification_channel AS ENUM ('sms', 'telegram', 'email');
CREATE TYPE client_notification_status AS ENUM ('queued', 'sent', 'failed', 'disabled', 'skipped');
CREATE TYPE client_notification_event AS ENUM (
  'device_received',
  'diagnostics_done',
  'repair_approval',
  'waiting_part',
  'ready_for_pickup',
  'pickup_reminder',
  'issued'
);
CREATE TYPE repair_approval_status AS ENUM ('pending', 'approved', 'declined', 'no_response', 'expired', 'manual_fixed');
CREATE TYPE payroll_rule_type AS ENUM ('fixed_monthly', 'per_unit', 'percent_from_work', 'fixed_plus_percent');
CREATE TYPE service_work_accrual_type AS ENUM ('percent_of_work_amount', 'fixed_per_unit');
CREATE TYPE payroll_period_status AS ENUM ('draft', 'calculated', 'approved', 'paid', 'cancelled');
CREATE TYPE vat_event_status AS ENUM ('act_issued', 'act_signed', 'payment_received', 'ready_for_tax_invoice', 'tax_invoice_registered', 'tax_invoice_overdue');
CREATE TYPE tax_invoice_status AS ENUM ('ready', 'sent', 'registered', 'error', 'stopped', 'overdue');
CREATE TYPE order_unit_status AS ENUM ('accepted', 'diagnostics', 'awaiting_approval', 'waiting_part', 'in_progress', 'ready', 'on_shelf', 'issued');

CREATE TABLE roles (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  name text NOT NULL,
  code text NOT NULL UNIQUE,
  description text
);

CREATE TABLE permissions (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  code text NOT NULL UNIQUE,
  name text NOT NULL,
  module text NOT NULL,
  description text
);

CREATE TABLE role_permissions (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  role_id uuid NOT NULL REFERENCES roles(id) ON DELETE CASCADE,
  permission_id uuid NOT NULL REFERENCES permissions(id) ON DELETE CASCADE,
  UNIQUE (role_id, permission_id)
);

CREATE TABLE users (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  full_name text NOT NULL,
  phone text,
  email text,
  login text NOT NULL UNIQUE,
  password_hash text NOT NULL,
  role_id uuid NOT NULL REFERENCES roles(id),
  is_active boolean NOT NULL DEFAULT true,
  last_login_at timestamptz,
  force_logout_at timestamptz,
  two_factor_enabled boolean NOT NULL DEFAULT false,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),
  deleted_at timestamptz
);

CREATE TRIGGER users_set_updated_at BEFORE UPDATE ON users
FOR EACH ROW EXECUTE FUNCTION set_updated_at();

CREATE TABLE clients (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  type client_type NOT NULL,
  full_name text,
  company_name text,
  phone text,
  email text,
  address text,
  notes text,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),
  deleted_at timestamptz,
  CHECK (
    (type = 'person' AND full_name IS NOT NULL)
    OR
    (type = 'company' AND company_name IS NOT NULL)
  )
);

CREATE TRIGGER clients_set_updated_at BEFORE UPDATE ON clients
FOR EACH ROW EXECUTE FUNCTION set_updated_at();

CREATE TABLE suppliers (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  name text NOT NULL,
  phone text,
  email text,
  address text,
  notes text,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),
  deleted_at timestamptz
);

CREATE TRIGGER suppliers_set_updated_at BEFORE UPDATE ON suppliers
FOR EACH ROW EXECUTE FUNCTION set_updated_at();

CREATE TABLE warehouses (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  name text NOT NULL,
  code text NOT NULL UNIQUE,
  address text,
  is_active boolean NOT NULL DEFAULT true
);

CREATE TABLE product_categories (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  name text NOT NULL,
  parent_id uuid REFERENCES product_categories(id)
);

CREATE TABLE products (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  sku text NOT NULL UNIQUE,
  article text,
  name text NOT NULL,
  category_id uuid REFERENCES product_categories(id),
  unit text NOT NULL DEFAULT 'pcs',
  barcode text UNIQUE,
  description text,
  default_purchase_price numeric(14,2) NOT NULL DEFAULT 0,
  default_sale_price numeric(14,2) NOT NULL DEFAULT 0,
  min_stock numeric(14,3) NOT NULL DEFAULT 0,
  is_service_part boolean NOT NULL DEFAULT true,
  is_active boolean NOT NULL DEFAULT true,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),
  deleted_at timestamptz
);

CREATE TRIGGER products_set_updated_at BEFORE UPDATE ON products
FOR EACH ROW EXECUTE FUNCTION set_updated_at();

CREATE TABLE product_locations (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  warehouse_id uuid NOT NULL REFERENCES warehouses(id),
  code text NOT NULL,
  name text NOT NULL,
  description text,
  UNIQUE (warehouse_id, code)
);

CREATE TABLE stock_balances (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  warehouse_id uuid NOT NULL REFERENCES warehouses(id),
  product_id uuid NOT NULL REFERENCES products(id),
  location_id uuid REFERENCES product_locations(id),
  qty_on_hand numeric(14,3) NOT NULL DEFAULT 0,
  qty_reserved numeric(14,3) NOT NULL DEFAULT 0,
  qty_available numeric(14,3) GENERATED ALWAYS AS (qty_on_hand - qty_reserved) STORED,
  updated_at timestamptz NOT NULL DEFAULT now(),
  UNIQUE (warehouse_id, product_id, location_id),
  CHECK (qty_on_hand >= 0),
  CHECK (qty_reserved >= 0),
  CHECK (qty_reserved <= qty_on_hand)
);

CREATE TABLE stock_movements (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  movement_type stock_movement_type NOT NULL,
  warehouse_id uuid NOT NULL REFERENCES warehouses(id),
  product_id uuid NOT NULL REFERENCES products(id),
  location_id uuid REFERENCES product_locations(id),
  qty numeric(14,3) NOT NULL,
  unit_cost numeric(14,2) NOT NULL DEFAULT 0,
  total_cost numeric(14,2) NOT NULL DEFAULT 0,
  reference_type text NOT NULL,
  reference_id uuid NOT NULL,
  comment text,
  created_by uuid REFERENCES users(id),
  created_at timestamptz NOT NULL DEFAULT now(),
  CHECK (qty <> 0)
);

CREATE TABLE inventory_sessions (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  warehouse_id uuid NOT NULL REFERENCES warehouses(id),
  started_by uuid NOT NULL REFERENCES users(id),
  started_at timestamptz NOT NULL DEFAULT now(),
  finished_at timestamptz,
  status inventory_session_status NOT NULL DEFAULT 'draft',
  notes text
);

CREATE TABLE inventory_session_items (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  inventory_session_id uuid NOT NULL REFERENCES inventory_sessions(id) ON DELETE CASCADE,
  product_id uuid NOT NULL REFERENCES products(id),
  expected_qty numeric(14,3) NOT NULL DEFAULT 0,
  actual_qty numeric(14,3) NOT NULL DEFAULT 0,
  difference_qty numeric(14,3) GENERATED ALWAYS AS (actual_qty - expected_qty) STORED,
  comment text,
  UNIQUE (inventory_session_id, product_id)
);

CREATE TABLE service_orders (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  order_number text NOT NULL UNIQUE,
  order_group_id uuid,
  qr_token text NOT NULL UNIQUE DEFAULT encode(gen_random_bytes(12), 'hex'),
  qr_url text,
  client_id uuid NOT NULL REFERENCES clients(id),
  device_type text NOT NULL,
  brand text,
  model text,
  serial_number text,
  problem_description text NOT NULL,
  equipment_description text,
  visual_condition text,
  manager_id uuid REFERENCES users(id),
  engineer_id uuid REFERENCES users(id),
  engineer_assigned_by uuid REFERENCES users(id),
  engineer_assigned_at timestamptz,
  engineer_accepted_at timestamptz,
  status service_order_status NOT NULL DEFAULT 'accepted',
  payment_status repair_payment_status NOT NULL DEFAULT 'unpaid',
  vat_status vat_event_status,
  act_issued_at timestamptz,
  act_signed_at timestamptz,
  debt_since timestamptz,
  status_changed_at timestamptz NOT NULL DEFAULT now(),
  intake_date timestamptz NOT NULL DEFAULT now(),
  promised_date timestamptz,
  diagnosis_result text,
  approval_date timestamptz,
  approved_by uuid REFERENCES users(id),
  approved_amount numeric(14,2) NOT NULL DEFAULT 0,
  test_result text,
  repair_comment text,
  issued_at timestamptz,
  issued_by uuid REFERENCES users(id),
  total_work_amount numeric(14,2) NOT NULL DEFAULT 0,
  total_parts_amount numeric(14,2) NOT NULL DEFAULT 0,
  delivery_amount numeric(14,2) NOT NULL DEFAULT 0,
  total_amount numeric(14,2) NOT NULL DEFAULT 0,
  prepayment_amount numeric(14,2) NOT NULL DEFAULT 0,
  debt_amount numeric(14,2) NOT NULL DEFAULT 0,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),
  deleted_at timestamptz
);

CREATE TRIGGER service_orders_set_updated_at BEFORE UPDATE ON service_orders
FOR EACH ROW EXECUTE FUNCTION set_updated_at();

CREATE TABLE service_order_status_history (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  service_order_id uuid NOT NULL REFERENCES service_orders(id) ON DELETE CASCADE,
  old_status service_order_status,
  new_status service_order_status NOT NULL,
  changed_by uuid REFERENCES users(id),
  changed_at timestamptz NOT NULL DEFAULT now(),
  comment text
);

CREATE TABLE service_order_assignments (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  service_order_id uuid NOT NULL REFERENCES service_orders(id) ON DELETE CASCADE,
  old_engineer_id uuid REFERENCES users(id),
  new_engineer_id uuid NOT NULL REFERENCES users(id),
  assigned_by uuid REFERENCES users(id),
  assigned_at timestamptz NOT NULL DEFAULT now(),
  accepted_at timestamptz,
  comment text
);

CREATE TABLE service_order_items (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  service_order_id uuid NOT NULL REFERENCES service_orders(id) ON DELETE CASCADE,
  work_name text NOT NULL,
  service_type text NOT NULL DEFAULT 'repair',
  engineer_id uuid REFERENCES users(id),
  qty numeric(14,3) NOT NULL DEFAULT 1,
  unit_price numeric(14,2) NOT NULL DEFAULT 0,
  total_price numeric(14,2) NOT NULL DEFAULT 0,
  payroll_accrual_type service_work_accrual_type NOT NULL DEFAULT 'percent_of_work_amount',
  payroll_percent_rate numeric(5,2) NOT NULL DEFAULT 0,
  payroll_fixed_per_unit numeric(14,2) NOT NULL DEFAULT 0,
  payroll_accrual_amount numeric(14,2) NOT NULL DEFAULT 0,
  payroll_locked_at timestamptz,
  comment text,
  CHECK (qty > 0),
  CHECK (payroll_percent_rate >= 0),
  CHECK (payroll_fixed_per_unit >= 0),
  CHECK (payroll_accrual_amount >= 0)
);

CREATE TABLE service_part_requirements (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  service_order_id uuid NOT NULL REFERENCES service_orders(id) ON DELETE CASCADE,
  product_id uuid NOT NULL REFERENCES products(id),
  qty_required numeric(14,3) NOT NULL DEFAULT 0,
  qty_reserved numeric(14,3) NOT NULL DEFAULT 0,
  qty_installed numeric(14,3) NOT NULL DEFAULT 0,
  status service_part_status NOT NULL DEFAULT 'required',
  requested_by uuid REFERENCES users(id),
  requested_at timestamptz NOT NULL DEFAULT now(),
  expected_date date,
  comment text,
  CHECK (qty_required > 0),
  CHECK (qty_reserved >= 0),
  CHECK (qty_installed >= 0)
);

CREATE TABLE service_parts_usage (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  service_order_id uuid NOT NULL REFERENCES service_orders(id) ON DELETE CASCADE,
  product_id uuid NOT NULL REFERENCES products(id),
  qty numeric(14,3) NOT NULL,
  unit_cost numeric(14,2) NOT NULL DEFAULT 0,
  total_cost numeric(14,2) NOT NULL DEFAULT 0,
  installed_by uuid REFERENCES users(id),
  installed_at timestamptz NOT NULL DEFAULT now(),
  comment text,
  CHECK (qty > 0)
);

CREATE TABLE purchase_orders (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  order_number text NOT NULL UNIQUE,
  supplier_id uuid NOT NULL REFERENCES suppliers(id),
  status purchase_order_status NOT NULL DEFAULT 'draft',
  order_date date NOT NULL DEFAULT CURRENT_DATE,
  expected_date date,
  created_by uuid REFERENCES users(id),
  approved_by uuid REFERENCES users(id),
  comment text,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),
  deleted_at timestamptz
);

CREATE TRIGGER purchase_orders_set_updated_at BEFORE UPDATE ON purchase_orders
FOR EACH ROW EXECUTE FUNCTION set_updated_at();

CREATE TABLE purchase_order_items (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  purchase_order_id uuid NOT NULL REFERENCES purchase_orders(id) ON DELETE CASCADE,
  product_id uuid NOT NULL REFERENCES products(id),
  qty numeric(14,3) NOT NULL,
  purchase_price numeric(14,2) NOT NULL DEFAULT 0,
  purchase_price_without_vat numeric(14,2) NOT NULL DEFAULT 0,
  vat_rate numeric(5,2) NOT NULL DEFAULT 20.00,
  vat_amount numeric(14,2) NOT NULL DEFAULT 0,
  total_price numeric(14,2) NOT NULL DEFAULT 0,
  linked_service_order_id uuid REFERENCES service_orders(id),
  linked_requirement_id uuid REFERENCES service_part_requirements(id),
  comment text,
  CHECK (qty > 0)
);

CREATE TABLE goods_receipts (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  receipt_number text NOT NULL UNIQUE,
  supplier_id uuid NOT NULL REFERENCES suppliers(id),
  purchase_order_id uuid REFERENCES purchase_orders(id),
  warehouse_id uuid NOT NULL REFERENCES warehouses(id),
  received_by uuid REFERENCES users(id),
  receipt_date timestamptz NOT NULL DEFAULT now(),
  status goods_receipt_status NOT NULL DEFAULT 'draft',
  comment text,
  created_at timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE goods_receipt_items (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  goods_receipt_id uuid NOT NULL REFERENCES goods_receipts(id) ON DELETE CASCADE,
  product_id uuid NOT NULL REFERENCES products(id),
  qty numeric(14,3) NOT NULL,
  purchase_price numeric(14,2) NOT NULL DEFAULT 0,
  purchase_price_without_vat numeric(14,2) NOT NULL DEFAULT 0,
  vat_rate numeric(5,2) NOT NULL DEFAULT 20.00,
  vat_amount numeric(14,2) NOT NULL DEFAULT 0,
  total_price numeric(14,2) NOT NULL DEFAULT 0,
  linked_service_order_id uuid REFERENCES service_orders(id),
  linked_requirement_id uuid REFERENCES service_part_requirements(id),
  comment text,
  CHECK (qty > 0)
);

CREATE TABLE sales (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  sale_number text NOT NULL UNIQUE,
  client_id uuid NOT NULL REFERENCES clients(id),
  manager_id uuid REFERENCES users(id),
  warehouse_id uuid NOT NULL REFERENCES warehouses(id),
  status sale_status NOT NULL DEFAULT 'draft',
  sale_date timestamptz NOT NULL DEFAULT now(),
  linked_service_order_id uuid REFERENCES service_orders(id),
  subtotal_amount numeric(14,2) NOT NULL DEFAULT 0,
  discount_amount numeric(14,2) NOT NULL DEFAULT 0,
  amount_without_vat numeric(14,2) NOT NULL DEFAULT 0,
  vat_rate numeric(5,2) NOT NULL DEFAULT 20.00,
  vat_amount numeric(14,2) NOT NULL DEFAULT 0,
  total_amount numeric(14,2) NOT NULL DEFAULT 0,
  paid_amount numeric(14,2) NOT NULL DEFAULT 0,
  debt_amount numeric(14,2) NOT NULL DEFAULT 0,
  comment text,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),
  deleted_at timestamptz
);

CREATE TRIGGER sales_set_updated_at BEFORE UPDATE ON sales
FOR EACH ROW EXECUTE FUNCTION set_updated_at();

CREATE TABLE sale_items (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  sale_id uuid NOT NULL REFERENCES sales(id) ON DELETE CASCADE,
  product_id uuid NOT NULL REFERENCES products(id),
  qty numeric(14,3) NOT NULL,
  sale_price numeric(14,2) NOT NULL DEFAULT 0,
  sale_price_without_vat numeric(14,2) NOT NULL DEFAULT 0,
  vat_rate numeric(5,2) NOT NULL DEFAULT 20.00,
  vat_amount numeric(14,2) NOT NULL DEFAULT 0,
  discount_amount numeric(14,2) NOT NULL DEFAULT 0,
  total_amount numeric(14,2) NOT NULL DEFAULT 0,
  unit_cost numeric(14,2) NOT NULL DEFAULT 0,
  total_cost numeric(14,2) NOT NULL DEFAULT 0,
  comment text,
  CHECK (qty > 0)
);

CREATE TABLE sale_returns (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  sale_id uuid NOT NULL REFERENCES sales(id),
  return_number text NOT NULL UNIQUE,
  return_date timestamptz NOT NULL DEFAULT now(),
  created_by uuid REFERENCES users(id),
  total_amount numeric(14,2) NOT NULL DEFAULT 0,
  total_without_vat numeric(14,2) NOT NULL DEFAULT 0,
  vat_amount numeric(14,2) NOT NULL DEFAULT 0,
  refund_method payment_method NOT NULL,
  comment text
);

CREATE TABLE sale_return_items (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  sale_return_id uuid NOT NULL REFERENCES sale_returns(id) ON DELETE CASCADE,
  product_id uuid NOT NULL REFERENCES products(id),
  qty numeric(14,3) NOT NULL,
  refund_amount numeric(14,2) NOT NULL DEFAULT 0,
  return_reason text NOT NULL,
  return_to_stock boolean NOT NULL DEFAULT true,
  write_off_as_defect boolean NOT NULL DEFAULT false,
  CHECK (qty > 0),
  CHECK (NOT (return_to_stock AND write_off_as_defect))
);

CREATE TABLE payments (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  payment_date timestamptz NOT NULL DEFAULT now(),
  client_id uuid NOT NULL REFERENCES clients(id),
  source_type payment_source_type NOT NULL,
  source_id uuid NOT NULL,
  payment_method payment_method NOT NULL,
  payment_kind payment_kind NOT NULL,
  amount numeric(14,2) NOT NULL,
  transaction_number text,
  accepted_by uuid REFERENCES users(id),
  comment text,
  created_at timestamptz NOT NULL DEFAULT now(),
  CHECK (amount <> 0)
);

CREATE TABLE cash_sessions (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  opened_by uuid NOT NULL REFERENCES users(id),
  opened_at timestamptz NOT NULL DEFAULT now(),
  closed_by uuid REFERENCES users(id),
  closed_at timestamptz,
  opening_cash_amount numeric(14,2) NOT NULL DEFAULT 0,
  opening_card_amount numeric(14,2) NOT NULL DEFAULT 0,
  closing_cash_amount numeric(14,2),
  closing_card_amount numeric(14,2),
  closing_bank_amount numeric(14,2),
  status cash_session_status NOT NULL DEFAULT 'open'
);

CREATE TABLE terminal_transactions (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  payment_id uuid REFERENCES payments(id),
  terminal_provider text,
  terminal_id text,
  transaction_number text NOT NULL,
  amount numeric(14,2) NOT NULL,
  status text NOT NULL DEFAULT 'manual_confirmed',
  raw_response_json jsonb,
  created_at timestamptz NOT NULL DEFAULT now(),
  CHECK (amount > 0)
);

CREATE TABLE cash_movements (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  cash_session_id uuid NOT NULL REFERENCES cash_sessions(id),
  movement_type cash_movement_type NOT NULL,
  amount numeric(14,2) NOT NULL,
  payment_id uuid REFERENCES payments(id),
  created_by uuid REFERENCES users(id),
  created_at timestamptz NOT NULL DEFAULT now(),
  comment text,
  CHECK (amount > 0)
);

CREATE TABLE product_prices (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  product_id uuid NOT NULL REFERENCES products(id) ON DELETE CASCADE,
  price_type price_type NOT NULL,
  price numeric(14,2) NOT NULL,
  start_date date NOT NULL DEFAULT CURRENT_DATE,
  end_date date,
  created_by uuid REFERENCES users(id),
  CHECK (price >= 0),
  CHECK (end_date IS NULL OR end_date >= start_date)
);

CREATE TABLE attachments (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  file_name text NOT NULL,
  file_path text NOT NULL,
  mime_type text NOT NULL,
  file_size bigint NOT NULL,
  entity_type entity_type NOT NULL,
  entity_id uuid NOT NULL,
  uploaded_by uuid REFERENCES users(id),
  uploaded_at timestamptz NOT NULL DEFAULT now(),
  CHECK (file_size >= 0)
);

CREATE TABLE document_templates (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  kind document_kind NOT NULL UNIQUE,
  name text NOT NULL,
  body_html text NOT NULL,
  company_name text NOT NULL DEFAULT 'НВКПП «СПЕКТР-АС»',
  company_edrpou text NOT NULL DEFAULT '22590485',
  logo_file_path text,
  is_active boolean NOT NULL DEFAULT true,
  version integer NOT NULL DEFAULT 1,
  created_by uuid REFERENCES users(id),
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),
  CHECK (version > 0)
);

CREATE TABLE generated_documents (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  document_number text NOT NULL UNIQUE,
  kind document_kind NOT NULL,
  status document_status NOT NULL DEFAULT 'draft',
  entity_type document_entity_type NOT NULL,
  entity_id uuid NOT NULL,
  client_id uuid REFERENCES clients(id),
  supplier_id uuid REFERENCES suppliers(id),
  payment_id uuid REFERENCES payments(id),
  template_id uuid REFERENCES document_templates(id),
  pdf_file_path text,
  version integer NOT NULL DEFAULT 1,
  created_by uuid REFERENCES users(id),
  signed_by uuid REFERENCES users(id),
  signed_at timestamptz,
  locked_at timestamptz,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),
  deleted_at timestamptz,
  CHECK (version > 0),
  CHECK ((status <> 'signed') OR (signed_by IS NOT NULL AND signed_at IS NOT NULL AND locked_at IS NOT NULL)),
  CHECK ((client_id IS NOT NULL) OR (supplier_id IS NOT NULL) OR (payment_id IS NOT NULL))
);

CREATE TABLE generated_document_versions (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  document_id uuid NOT NULL REFERENCES generated_documents(id) ON DELETE CASCADE,
  version integer NOT NULL,
  pdf_file_path text NOT NULL,
  change_reason text NOT NULL,
  created_by uuid REFERENCES users(id),
  created_at timestamptz NOT NULL DEFAULT now(),
  UNIQUE (document_id, version),
  CHECK (version > 0)
);

CREATE TABLE audit_logs (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id uuid REFERENCES users(id),
  action_type audit_action_type NOT NULL,
  entity_type text NOT NULL,
  entity_id uuid,
  old_value_json jsonb,
  new_value_json jsonb,
  ip_address inet,
  user_agent text,
  created_at timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE notifications (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id uuid NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  title text NOT NULL,
  body text NOT NULL,
  type notification_type NOT NULL,
  is_read boolean NOT NULL DEFAULT false,
  created_at timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE client_notification_templates (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  event client_notification_event NOT NULL,
  channel client_notification_channel NOT NULL,
  name text NOT NULL,
  message_template text NOT NULL,
  is_enabled boolean NOT NULL DEFAULT true,
  created_by uuid REFERENCES users(id),
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),
  UNIQUE (event, channel)
);

CREATE TABLE client_notifications (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  service_order_id uuid NOT NULL REFERENCES service_orders(id) ON DELETE CASCADE,
  client_id uuid REFERENCES clients(id),
  template_id uuid REFERENCES client_notification_templates(id),
  channel client_notification_channel NOT NULL,
  event client_notification_event NOT NULL,
  status client_notification_status NOT NULL DEFAULT 'queued',
  recipient text NOT NULL,
  message_text text NOT NULL,
  provider_message_id text,
  error_text text,
  created_by uuid REFERENCES users(id),
  created_at timestamptz NOT NULL DEFAULT now(),
  sent_at timestamptz,
  CHECK (recipient <> ''),
  CHECK (message_text <> '')
);

CREATE TABLE repair_approvals (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  service_order_id uuid NOT NULL REFERENCES service_orders(id) ON DELETE CASCADE,
  approval_token text NOT NULL UNIQUE,
  status repair_approval_status NOT NULL DEFAULT 'pending',
  client_id uuid REFERENCES clients(id),
  client_phone text NOT NULL,
  device_snapshot text NOT NULL,
  issue_snapshot text NOT NULL,
  works_snapshot_json jsonb NOT NULL DEFAULT '[]'::jsonb,
  parts_snapshot_json jsonb NOT NULL DEFAULT '[]'::jsonb,
  total_amount numeric(14,2) NOT NULL,
  promised_date_snapshot timestamptz,
  message_text text NOT NULL,
  sent_channel client_notification_channel NOT NULL,
  sent_at timestamptz NOT NULL DEFAULT now(),
  response_at timestamptz,
  response_ip inet,
  response_user_agent text,
  response_comment text,
  created_by uuid REFERENCES users(id),
  created_at timestamptz NOT NULL DEFAULT now(),
  CHECK (total_amount >= 0),
  CHECK (client_phone <> ''),
  CHECK (message_text <> '')
);

CREATE TABLE order_groups (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  client_id uuid NOT NULL REFERENCES clients(id),
  intake_date timestamptz NOT NULL DEFAULT now(),
  comment text,
  created_by uuid REFERENCES users(id),
  created_at timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE storage_cells (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  shelf text NOT NULL,
  level integer NOT NULL,
  position integer NOT NULL,
  code text NOT NULL UNIQUE,
  is_active boolean NOT NULL DEFAULT true,
  CHECK (level > 0),
  CHECK (position > 0)
);

CREATE TABLE order_units (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  service_order_id uuid NOT NULL REFERENCES service_orders(id) ON DELETE CASCADE,
  order_group_id uuid REFERENCES order_groups(id),
  unit_code text NOT NULL UNIQUE,
  unit_type text NOT NULL,
  qr_url text,
  barcode text,
  status order_unit_status NOT NULL DEFAULT 'accepted',
  engineer_id uuid REFERENCES users(id),
  storage_cell_id uuid REFERENCES storage_cells(id),
  accepted_at timestamptz NOT NULL DEFAULT now(),
  issued_at timestamptz,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE order_unit_movements (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  order_unit_id uuid NOT NULL REFERENCES order_units(id) ON DELETE CASCADE,
  old_status order_unit_status,
  new_status order_unit_status NOT NULL,
  storage_cell_id uuid REFERENCES storage_cells(id),
  changed_by uuid REFERENCES users(id),
  changed_at timestamptz NOT NULL DEFAULT now(),
  comment text
);

CREATE TABLE tax_invoices (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  tax_invoice_number text UNIQUE,
  service_order_id uuid NOT NULL REFERENCES service_orders(id),
  client_id uuid NOT NULL REFERENCES clients(id),
  event_date date NOT NULL,
  registration_deadline date NOT NULL,
  amount numeric(14,2) NOT NULL DEFAULT 0,
  vat_amount numeric(14,2) NOT NULL DEFAULT 0,
  status tax_invoice_status NOT NULL DEFAULT 'ready',
  formed_at timestamptz NOT NULL DEFAULT now(),
  sent_at timestamptz,
  registered_at timestamptz,
  receipt_json jsonb,
  response_text text,
  created_by uuid REFERENCES users(id),
  updated_at timestamptz NOT NULL DEFAULT now(),
  CHECK (amount >= 0),
  CHECK (vat_amount >= 0),
  CHECK (registration_deadline >= event_date)
);

CREATE TABLE receivables_control (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  client_id uuid NOT NULL REFERENCES clients(id),
  service_order_id uuid REFERENCES service_orders(id),
  debt_amount numeric(14,2) NOT NULL DEFAULT 0,
  debt_since date,
  act_issued_at date,
  act_signed_at date,
  last_contact_at timestamptz,
  responsible_manager_id uuid REFERENCES users(id),
  escalation_at timestamptz,
  status text NOT NULL DEFAULT 'active',
  created_at timestamptz NOT NULL DEFAULT now(),
  CHECK (debt_amount >= 0)
);

CREATE TABLE payroll_rules (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id uuid NOT NULL REFERENCES users(id),
  rule_type payroll_rule_type NOT NULL,
  monthly_rate numeric(14,2) NOT NULL DEFAULT 0,
  per_unit_rate numeric(14,2) NOT NULL DEFAULT 0,
  percent_rate numeric(5,2) NOT NULL DEFAULT 0,
  bonus_amount numeric(14,2) NOT NULL DEFAULT 0,
  penalty_per_return numeric(14,2) NOT NULL DEFAULT 0,
  is_active boolean NOT NULL DEFAULT true,
  start_date date NOT NULL DEFAULT CURRENT_DATE,
  end_date date,
  created_by uuid REFERENCES users(id),
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),
  CHECK (monthly_rate >= 0),
  CHECK (per_unit_rate >= 0),
  CHECK (percent_rate >= 0),
  CHECK (bonus_amount >= 0),
  CHECK (penalty_per_return >= 0),
  CHECK (end_date IS NULL OR end_date >= start_date)
);

CREATE TABLE payroll_periods (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  period_start date NOT NULL,
  period_end date NOT NULL,
  status payroll_period_status NOT NULL DEFAULT 'draft',
  created_by uuid REFERENCES users(id),
  approved_by uuid REFERENCES users(id),
  paid_by uuid REFERENCES users(id),
  created_at timestamptz NOT NULL DEFAULT now(),
  approved_at timestamptz,
  paid_at timestamptz,
  CHECK (period_end >= period_start)
);

CREATE TABLE payroll_accruals (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  payroll_period_id uuid NOT NULL REFERENCES payroll_periods(id) ON DELETE CASCADE,
  user_id uuid NOT NULL REFERENCES users(id),
  service_order_id uuid REFERENCES service_orders(id),
  service_order_item_id uuid REFERENCES service_order_items(id),
  payroll_rule_id uuid REFERENCES payroll_rules(id),
  accrual_type service_work_accrual_type NOT NULL,
  qty numeric(14,3) NOT NULL DEFAULT 1,
  work_amount numeric(14,2) NOT NULL DEFAULT 0,
  percent_rate numeric(5,2) NOT NULL DEFAULT 0,
  fixed_per_unit numeric(14,2) NOT NULL DEFAULT 0,
  accrual_amount numeric(14,2) NOT NULL DEFAULT 0,
  bonus_amount numeric(14,2) NOT NULL DEFAULT 0,
  penalty_amount numeric(14,2) NOT NULL DEFAULT 0,
  comment text,
  created_at timestamptz NOT NULL DEFAULT now(),
  CHECK (qty > 0),
  CHECK (work_amount >= 0),
  CHECK (percent_rate >= 0),
  CHECK (fixed_per_unit >= 0),
  CHECK (accrual_amount >= 0),
  CHECK (bonus_amount >= 0),
  CHECK (penalty_amount >= 0)
);

CREATE INDEX idx_users_role_id ON users(role_id);
CREATE INDEX idx_clients_phone ON clients(phone);
CREATE INDEX idx_clients_type ON clients(type);
CREATE INDEX idx_products_category_id ON products(category_id);
CREATE INDEX idx_products_barcode ON products(barcode);
CREATE INDEX idx_products_name ON products USING gin (to_tsvector('simple', name));
CREATE INDEX idx_stock_balances_product_id ON stock_balances(product_id);
CREATE INDEX idx_stock_balances_warehouse_product ON stock_balances(warehouse_id, product_id);
CREATE INDEX idx_stock_movements_product_id ON stock_movements(product_id);
CREATE INDEX idx_stock_movements_created_at ON stock_movements(created_at);
CREATE INDEX idx_stock_movements_reference ON stock_movements(reference_type, reference_id);
CREATE INDEX idx_service_orders_client_id ON service_orders(client_id);
CREATE INDEX idx_service_orders_qr_token ON service_orders(qr_token);
CREATE INDEX idx_service_orders_status ON service_orders(status);
CREATE INDEX idx_service_orders_payment_status ON service_orders(payment_status);
CREATE INDEX idx_service_orders_vat_status ON service_orders(vat_status);
CREATE INDEX idx_service_orders_order_group_id ON service_orders(order_group_id);
CREATE INDEX idx_service_orders_status_changed_at ON service_orders(status_changed_at);
CREATE INDEX idx_service_orders_intake_date ON service_orders(intake_date);
CREATE INDEX idx_service_orders_engineer_id ON service_orders(engineer_id);
CREATE INDEX idx_service_orders_engineer_accepted_at ON service_orders(engineer_accepted_at);
CREATE INDEX idx_service_order_assignments_order_id ON service_order_assignments(service_order_id);
CREATE INDEX idx_service_order_assignments_new_engineer_id ON service_order_assignments(new_engineer_id);
CREATE INDEX idx_service_order_items_engineer_id ON service_order_items(engineer_id);
CREATE INDEX idx_service_order_items_accrual_type ON service_order_items(payroll_accrual_type);
CREATE INDEX idx_service_part_requirements_status ON service_part_requirements(status);
CREATE INDEX idx_service_part_requirements_product_id ON service_part_requirements(product_id);
CREATE INDEX idx_purchase_orders_supplier_id ON purchase_orders(supplier_id);
CREATE INDEX idx_purchase_orders_status ON purchase_orders(status);
CREATE INDEX idx_purchase_orders_order_date ON purchase_orders(order_date);
CREATE INDEX idx_purchase_order_items_product_id ON purchase_order_items(product_id);
CREATE INDEX idx_goods_receipts_purchase_order_id ON goods_receipts(purchase_order_id);
CREATE INDEX idx_goods_receipts_receipt_date ON goods_receipts(receipt_date);
CREATE INDEX idx_sales_client_id ON sales(client_id);
CREATE INDEX idx_sales_manager_id ON sales(manager_id);
CREATE INDEX idx_sales_status ON sales(status);
CREATE INDEX idx_sales_sale_date ON sales(sale_date);
CREATE INDEX idx_sales_linked_service_order_id ON sales(linked_service_order_id);
CREATE INDEX idx_sales_vat_rate ON sales(vat_rate);
CREATE INDEX idx_sale_items_product_id ON sale_items(product_id);
CREATE INDEX idx_payments_client_id ON payments(client_id);
CREATE INDEX idx_payments_source ON payments(source_type, source_id);
CREATE INDEX idx_payments_payment_date ON payments(payment_date);
CREATE INDEX idx_payments_method ON payments(payment_method);
CREATE INDEX idx_payments_transaction_number ON payments(transaction_number);
CREATE INDEX idx_cash_movements_session_id ON cash_movements(cash_session_id);
CREATE INDEX idx_terminal_transactions_payment_id ON terminal_transactions(payment_id);
CREATE INDEX idx_terminal_transactions_number ON terminal_transactions(transaction_number);
CREATE INDEX idx_product_prices_product_type ON product_prices(product_id, price_type);
CREATE INDEX idx_attachments_entity ON attachments(entity_type, entity_id);
CREATE INDEX idx_document_templates_kind ON document_templates(kind);
CREATE INDEX idx_generated_documents_entity ON generated_documents(entity_type, entity_id);
CREATE INDEX idx_generated_documents_kind ON generated_documents(kind);
CREATE INDEX idx_generated_documents_status ON generated_documents(status);
CREATE INDEX idx_generated_documents_created_at ON generated_documents(created_at);
CREATE INDEX idx_generated_document_versions_document_id ON generated_document_versions(document_id);
CREATE INDEX idx_audit_logs_user_id ON audit_logs(user_id);
CREATE INDEX idx_audit_logs_entity ON audit_logs(entity_type, entity_id);
CREATE INDEX idx_audit_logs_created_at ON audit_logs(created_at);
CREATE INDEX idx_notifications_user_read ON notifications(user_id, is_read);
CREATE INDEX idx_client_notification_templates_event_channel ON client_notification_templates(event, channel);
CREATE INDEX idx_client_notifications_order_id ON client_notifications(service_order_id);
CREATE INDEX idx_client_notifications_status ON client_notifications(status);
CREATE INDEX idx_client_notifications_created_at ON client_notifications(created_at);
CREATE INDEX idx_repair_approvals_order_id ON repair_approvals(service_order_id);
CREATE INDEX idx_repair_approvals_token ON repair_approvals(approval_token);
CREATE INDEX idx_repair_approvals_status ON repair_approvals(status);
CREATE INDEX idx_repair_approvals_sent_at ON repair_approvals(sent_at);
CREATE INDEX idx_order_units_order_id ON order_units(service_order_id);
CREATE INDEX idx_order_units_code ON order_units(unit_code);
CREATE INDEX idx_order_units_status ON order_units(status);
CREATE INDEX idx_order_units_storage_cell_id ON order_units(storage_cell_id);
CREATE INDEX idx_storage_cells_code ON storage_cells(code);
CREATE INDEX idx_tax_invoices_order_id ON tax_invoices(service_order_id);
CREATE INDEX idx_tax_invoices_status ON tax_invoices(status);
CREATE INDEX idx_tax_invoices_deadline ON tax_invoices(registration_deadline);
CREATE INDEX idx_receivables_control_client_id ON receivables_control(client_id);
CREATE INDEX idx_payroll_rules_user_id ON payroll_rules(user_id);
CREATE INDEX idx_payroll_rules_active ON payroll_rules(is_active);
CREATE INDEX idx_payroll_periods_dates ON payroll_periods(period_start, period_end);
CREATE INDEX idx_payroll_periods_status ON payroll_periods(status);
CREATE INDEX idx_payroll_accruals_period_id ON payroll_accruals(payroll_period_id);
CREATE INDEX idx_payroll_accruals_user_id ON payroll_accruals(user_id);
CREATE INDEX idx_payroll_accruals_order_item_id ON payroll_accruals(service_order_item_id);
