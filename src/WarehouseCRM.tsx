import React, { useMemo, useState, useEffect } from 'react';
import { Player } from '@remotion/player';
import * as XLSX from 'xlsx';
import { useRef } from 'react';
import { OrderStatus } from './types';
import invoiceTemplateHtml from './templates/invoiceTemplate.html?raw';
import actTemplateHtml from './templates/actTemplate.html?raw';
import waybillTemplateHtml from './templates/waybillTemplate.html?raw';
import { 
  Archive,
  Banknote,
  Bell,
  CheckCircle2,
  ClipboardList,
  History,
  LayoutDashboard,
  LogOut,
  Menu,
  PackageCheck,
  PackagePlus,
  Plus,
  Search,
  Settings,
  ShieldCheck,
  ShoppingCart,
  Truck,
  Users,
  Wrench,
  X,
} from 'lucide-react';
import { ServiceFlowAnimation } from './remotion/ServiceFlowAnimation';

type Page = 'dashboard' | 'inbox' | 'employee-control' | 'contracts' | 'bank-import' | 'orders' | 'my-orders' | 'sales' | 'clients' | 'parts' | 'purchases' | 'cash' | 'returns' | 'documents' | 'tax-invoices' | 'bas-exchange' | 'problem-clients' | 'order-units' | 'storage' | 'movements' | 'finance' | 'reports' | 'payroll' | 'acceptance' | 'team' | 'settings';
type PaymentMethod = 'Готівка' | 'Картка' | 'Безготівка' | 'Змішана';
type PaymentType = 'Передплата' | 'Часткова оплата' | 'Повна оплата' | 'Доплата' | 'Повернення коштів';
type SaleStatus = 'Чернетка' | 'Зарезервовано' | 'Очікує оплати' | 'Частково оплачено' | 'Оплачено' | 'Видано' | 'Закрито' | 'Скасовано' | 'Повернення';

type RepairPaymentStatus = 'Не оплачено' | 'Передплата' | 'Частково оплачено' | 'Повністю оплачено' | 'Повернення' | 'Борг';
type SimpleRepairPaymentStatus = 'Не оплачено' | 'Частично оплачено' | 'Оплачено';
type SimpleWorkType = 'Ремонт техники' | 'Заправка картриджа' | 'Регенерация картриджа';
type OrderControlIssue = { code: string; message: string };
type VatStatus = 'Акт видано' | 'Акт підписано' | 'Оплату отримано' | 'Очікує ПН' | 'ПН створено' | 'Готово до реєстрації ПН' | 'ПН зареєстрована' | 'Прострочена реєстрація ПН';
type TaxInvoiceStatus = 'Очікує ПН' | 'Створено' | 'Зареєстровано' | 'Помилка';
type TaxInvoiceEventType = 'payment' | 'act';
type OrderUnitStatus = 'Прийнято' | 'Діагностика' | 'Очікує погодження' | 'Жде запчастину' | 'В ремонті' | 'Готово' | 'На полці' | 'Видано';
type BasExchangeStatus = 'new' | 'queued' | 'sent' | 'imported' | 'error';
type BasObjectType = 'counterparty' | 'act' | 'sale' | 'payment' | 'purchase' | 'tax_invoice';

type StatusHistoryEntry = {
  id: string;
  oldStatus?: OrderStatus;
  newStatus: OrderStatus;
  changedBy: string;
  changedAt: string;
  comment: string;
};
type OrderPartStatus =
  | 'Не потрібно'
  | 'Потрібно'
  | 'До закупівлі'
  | 'Замовлено'
  | 'В дорозі'
  | 'Прибуло'
  | 'Зарезервовано'
  | 'Видано інженеру'
  | 'Встановлено'
  | 'Списано'
  | 'Повернення';
type RequirementStatus = 'Потрібно' | 'До закупівлі' | 'Замовлено' | 'В дорозі' | 'Прибуло';
type PurchaseStatus = 'Створено' | 'Нова' | 'В роботі' | 'Замовлено' | 'В дорозі' | 'Прибуло' | 'Отримано' | 'Частково прибуло';
type MovementType = 'Прихід' | 'Резерв' | 'Зняття резерву' | 'Видача інженеру' | 'Встановлення' | 'Списання' | 'Продаж' | 'Повернення' | 'Коригування';
type PurchaseReason = 'Мінімум' | 'Під замовлення' | 'Вручну';
type PurchasePriority = 'Низький' | 'Середній' | 'Високий';

type ProductBatch = {
  id: string;
  productId: string;
  qtyTotal: number;
  qtyAvailable: number;
  purchasePrice: number;
  purchaseDate: string;
  source: string;
  supplier?: string;
  documentNo?: string;
  comment?: string;
};

type BatchAllocation = {
  batchId: string;
  qty: number;
  unitCost: number;
  purchaseDate: string;
};

type Product = {
  id: string;
  name: string;
  sku: string;
  barcode: string;
  extraBarcodes: string[];
  category: string;
  brand: string;
  model: string;
  unit: string;
  cost: number;
  price: number;
  stock: number;
  reserved: number;
  withEngineer: number;
  installed: number;
  min: number;
  storageLocation: string;
  batches: ProductBatch[];
};

type ClientRecord = {
  name: string;
  phone: string;
  email?: string;
  taxId?: string;
  orders: number;
};

type ContractStatus = 'Активний' | 'Закритий';
type ContractActStatus = 'Не оплачено' | 'Частично оплачен' | 'Оплачен';
type SimpleLedgerEntityType = 'order' | 'contract' | 'act' | 'client';
type SimpleLedgerPaymentKind = 'предоплата' | 'оплата' | 'частичная';
type SimpleLedgerPaymentStatus = 'черновик' | 'ожидается' | 'проведено' | 'ожидает поступления' | 'подтвержден' | 'скасовано' | 'повернення';
type BankImportItemStatus = 'matched' | 'review' | 'unmatched';
type BankImportField = 'date' | 'amount' | 'currency' | 'payer' | 'taxId' | 'purpose' | 'documentRef' | 'direction' | 'bankAccount';
type BankDirection = 'incoming' | 'outgoing' | 'unknown';

type BankAccountRecord = {
  id: string;
  bankName: string;
  iban: string;
  currency: string;
  legalType: 'ФОП' | 'ТОВ';
};

type BankImportDraft = {
  fileName: string;
  headers: string[];
  rows: Record<string, unknown>[];
  detectedBank: string;
  mapping: Partial<Record<BankImportField, string>>;
};

type BankStatementRow = {
  id: string;
  date: string;
  amount: number;
  currency: string;
  payer: string;
  taxId: string;
  purpose: string;
  documentRef: string;
  direction: BankDirection;
  bankAccount: string;
};

type BankImportCandidate = {
  key: string;
  kind: 'order' | 'act' | 'invoice';
  entityId: string;
  client: string;
  amount: number;
  label: string;
  reference: string;
  taxId: string;
  eventDate?: string;
  backingEntityType?: PrintDocument['entityType'];
  backingEntityId?: string;
  documentId?: string;
};

type BankImportItem = {
  id: string;
  row: BankStatementRow;
  status: BankImportItemStatus;
  matchedCandidateKey?: string;
  matchedBy?: 'auto' | 'manual';
  reason: string;
  candidates: BankImportCandidate[];
};

type ContractRecord = {
  id: string;
  client: string;
  amount: number;
  startDate: string;
  endDate: string;
  status: ContractStatus;
  createdAt: string;
  closedAt?: string;
};

type ContractActRecord = {
  id: string;
  contractId: string;
  orderIds: string[];
  amount: number;
  paidAmount: number;
  remainingAmount: number;
  date: string;
  status: ContractActStatus;
};

type SimpleOrderPaymentRecord = {
  id: string;
  client: string;
  entityType: SimpleLedgerEntityType;
  entityId: string;
  clientTaxId?: string;
  orderId?: string;
  contractId?: string;
  actId?: string;
  amount: number;
  method: 'наличные' | 'карта' | 'перевод';
  paymentKind: SimpleLedgerPaymentKind;
  status: SimpleLedgerPaymentStatus;
  date: string;
  currency?: string;
  payer?: string;
  taxId?: string;
  purpose?: string;
  documentRef?: string;
  direction?: BankDirection;
  bankAccount?: string;
  bankName?: string;
  fingerprint?: string;
  acceptedBy?: string;
  cashShiftId?: string;
  countedAtShiftId?: string;
  cancelledAt?: string;
  cancelledBy?: string;
  refundReason?: string;
};

type OrderDebtLifecycleStatus = 'Новый долг' | 'В работе (частично оплачен)' | 'Закрыт';

type SimpleOrderActivityRecord = {
  id: string;
  date: string;
  action: string;
  detail?: string;
};

type OrderPart = {
  id: string;
  productId: string;
  qty: number;
  status: OrderPartStatus;
  cost: number;
  price: number;
  requirementId?: string;
  purchaseId?: string;
  batchAllocations?: BatchAllocation[];
};

type WorkAccrualType = 'percent_of_work_amount' | 'fixed_per_unit';
type ServiceWork = {
  name: string;
  serviceType: string;
  price: number;
  qty?: number;
  engineer?: string;
  accrualType: WorkAccrualType;
  payrollPercent?: number;
  payrollFixedPerUnit?: number;
};

type ServiceOrder = {
  id: string;
  groupId?: string;
  assignedTo?: string;
  createdByUserId?: string;
  createdAt?: string;
  issuedAt?: string;
  qrUrl: string;
  client: string;
  phone: string;
  device: string;
  serial: string;
  issue: string;
  appearance?: string;
  intakeComment?: string;
  estimatedAmount?: number;
  repairPrice?: number;
  repairPaymentMethod?: 'наличные' | 'карта' | 'перевод';
  contractId?: string;
  contractAccountedAt?: string;
  contractActId?: string;
  engineer: string;
  engineerAcceptedAt?: string;
  manager: string;
  legalEntity?: boolean;
  vatStatus?: VatStatus;
  actIssuedAt?: string;
  actReturnedAt?: string;
  debtSince?: string;
  issuedDebtAmount?: number;
  issuedDebtAt?: string;
  issuedDebtManager?: string;
  status: OrderStatus;
  statusChangedAt: string;
  intakeDate: string;
  promisedDate: string;
  diagnosisResult?: string;
  approvalComment?: string;
  testResult?: string;
  clientComment?: string;
  deliveryAmount?: number;
  additionalExpenses?: number;
  consumablesAmount?: number;
  externalServicesAmount?: number;
  works: ServiceWork[];
  parts: OrderPart[];
  payments: Payment[];
  locationCode?: string;
  locationStatus?: 'У комірці' | 'У інженера' | 'Видано';
  takenInRepairAt?: string;
  takenInRepairBy?: string;
  takenFromLocation?: string;
  returnedToCellAt?: string;
  returnedToCellBy?: string;
  urgent?: boolean;
  clientNotified?: boolean;
  lastNotificationType?: NotificationEvent;
  lastNotificationAt?: string;
  notificationHistory?: ClientNotification[];
  pendingExtraApproval?: boolean;
  engineerWorkCompletedAt?: string;
  engineerWorkCompletedBy?: string;
  currentVersion?: number;
  statusHistory: StatusHistoryEntry[];
  activityLog?: SimpleOrderActivityRecord[];
};

type OrderVersion = {
  id: string;
  orderId: string;
  versionNo: number;
  createdAt: string;
  createdBy: string;
  reason: string;
  snapshotData: string;
};

type PartRequirement = {
  id: string;
  orderId: string;
  productId: string;
  qty: number;
  status: RequirementStatus;
  reason?: PurchaseReason;
  priority?: PurchasePriority;
  requester?: string;
  comment?: string;
  supplier?: string;
  purchasePrice?: number;
  sourceLink?: string;
};

type PurchaseOrder = {
  id: string;
  supplier: string;
  items: Array<{ productId: string; qty: number; price: number; received: number; requirementId?: string; orderId?: string }>;
  status: PurchaseStatus;
  orderedAt: string;
  expectedAt: string;
  reason?: PurchaseReason;
  priority?: PurchasePriority;
  sourceLink?: string;
  comment?: string;
  paymentType?: 'bank' | 'cash';
  bankPaymentId?: string;
  cashAmount?: number;
  purchasedAt?: string;
  receivedAt?: string;
  requestedBy?: string;
};

type GoodsReceipt = {
  id: string;
  date: string;
  supplier: string;
  productId: string;
  qty: number;
  price: number;
  purchaseId: string;
};

type StockMovement = {
  id: string;
  date: string;
  type: MovementType;
  productId: string;
  qty: number;
  orderId?: string;
  purchaseId?: string;
  actor?: string;
  basis?: string;
  batchRefs?: string;
  unitPrice?: number;
  totalAmount?: number;
  comment: string;
};

type StockIntakeInput = {
  productId?: string;
  name: string;
  qty: number;
  purchasePrice: number;
  supplier?: string;
  category?: string;
  salePrice?: number;
  minStock?: number;
  barcode?: string;
  sku?: string;
  brand?: string;
  model?: string;
  unit?: string;
  storageLocation?: string;
  purchaseDocumentNo?: string;
  appendBarcodeToProductId?: string;
  forceCreateNew?: boolean;
};

type Payment = {
  id: string;
  date: string;
  amount: number;
  method: PaymentMethod;
  type: PaymentType;
  transactionNo: string;
  acceptedBy: string;
  status?: 'Чернетка' | 'Проведено' | 'Очікує надходження' | 'Підтверджено' | 'Очікує підтвердження' | 'Скасовано' | 'Повернення';
  confirmedBy?: string;
  confirmedAt?: string;
  saleId?: string;
  orderId?: string;
  comment: string;
  cashShiftId?: string;
  countedAtShiftId?: string;
  cancelledAt?: string;
  cancelledBy?: string;
  refundReason?: string;
};

type AccountingPeriod = 'today' | 'week' | 'month';
type AccountingExportKind = 'sales' | 'purchases' | 'stock' | 'summary';
type AccountingExportFormat = 'excel' | 'csv' | 'json';

type SaleItem = {
  id: string;
  productId: string;
  qty: number;
  price: number;
  cost: number;
  discount: number;
  status: 'Чернетка' | 'Зарезервовано' | 'Видано' | 'Повернення' | 'Скасовано';
  batchAllocations?: BatchAllocation[];
};

type SaleReturn = {
  id: string;
  saleId: string;
  productId: string;
  qty: number;
  reason: string;
  refund: number;
  destination: 'На склад' | 'Брак';
};

type Sale = {
  id: string;
  client: string;
  manager: string;
  date: string;
  status: SaleStatus;
  serviceOrderId?: string;
  items: SaleItem[];
  payments: Payment[];
  returns: SaleReturn[];
};

type TaxInvoice = {
  id: string;
  number: string;
  date: string;
  orderId: string;
  invoiceId?: string;
  actId?: string;
  paymentId?: string;
  waybillId?: string;
  eventType?: TaxInvoiceEventType;
  client: string;
  eventDate: string;
  amount: number;
  vatAmount: number;
  status: TaxInvoiceStatus;
  snapshot: string;
  createdAt: string;
  registeredAt?: string;
  responsible: string;
};

type TaxInvoiceSnapshot = {
  companyName: string;
  companyEdrpou: string;
  companyIpn: string;
  companyIban: string;
  companyBank: string;
  companyMfo: string;
  companyAddress: string;
  companyPhone: string;
  companyVatStatus: string;
  clientName: string;
  clientPhone: string;
  clientTaxId: string;
  orderId: string;
  orderNumber: string;
  device: string;
  managerName: string;
  documentNumber: string;
  documentDate: string;
  eventDate: string;
  eventType: TaxInvoiceEventType;
  sumWithoutVAT: number;
  vatAmount: number;
  sumWithVAT: number;
  totalWithVATWords: string;
  vatAmountWords: string;
  lines: Array<{ type: 'service' | 'product'; label: string; qty: number; price: number; total: number }>;
  paymentId?: string;
  paymentType?: string;
  paymentStatus?: string;
  invoiceId?: string;
  actId?: string;
  actNumber?: string;
  waybillId?: string;
  createdAt: string;
};

type BasExchangeItem = {
  id: string;
  platformEntityType: BasObjectType;
  platformEntityId: string;
  basEntityId?: string;
  exchangeStatus: BasExchangeStatus;
  payloadJson: string;
  sentAt?: string;
  confirmedAt?: string;
  retryCount: number;
  lastError?: string;
  amountNet: number;
  amountVat: number;
  amountGross: number;
  orderId?: string;
  client: string;
  createdBy: string;
  title: string;
  responsible: string;
};

type BasMapping = {
  directory: 'Контрагенти' | 'Номенклатура' | 'Послуги' | 'Склади' | 'Ставки ПДВ';
  platformValue: string;
  basValue: string;
  status: 'Зіставлено' | 'Потрібно зіставити';
};

type BasImplementationPhase = {
  phase: string;
  title: string;
  result: string;
  priority: 'Перша черга' | 'Друга черга' | 'Не робити на старті';
  items: string[];
  done: boolean;
};

type OrderUnit = {
  id: string;
  orderId: string;
  groupId?: string;
  client: string;
  type: 'Ноутбук' | 'Принтер' | 'Картридж' | 'МФУ';
  code: string;
  status: OrderUnitStatus;
  engineer: string;
  locationCode?: string;
  lastActionAt: string;
};

type WarehouseZone = 'REPAIR' | 'READY' | 'STORAGE' | 'OUT';
type WarehouseLocation = {
  id: string;
  zone: WarehouseZone;
  rack: string;
  shelf: string;
  code: string;
};

type OrderMovementLog = {
  id: string;
  orderId: string;
  fromLocation?: string;
  toLocation: string;
  userId: string;
  timestamp: string;
};

type StorageCell = {
  id: string;
  shelf: string;
  level: number;
  position: number;
  code: string;
  occupiedBy?: string;
};

type CashShift = {
  id: string;
  openedAt: string;
  openedBy: string;
  closedAt?: string;
  closedBy?: string;
  status: 'Відкрита' | 'Закрита';
  openingCash: number;
  cashIncome: number;
  cardIncome: number;
  bankIncome: number;
  cashExpense: number;
  actualCash?: number;
  difference?: number;
  closeComment?: string;
};

type DocumentKind =
  | 'Квитанція прийому'
  | 'Акт діагностики'
  | 'Заказ-наряд'
  | 'Рахунок на оплату'
  | 'Акт надання послуг'
  | 'Акт виконаних робіт'
  | 'Акт видачі'
  | 'Видаткова накладна'
  | 'Наряд на ремонт'
  | 'Швидкий наряд'
  | 'Акт технічного стану'
  | 'Гарантійний талон'
  | 'Товарний чек'
  | 'Накладна продажу'
  | 'Касовий чек'
  | 'Документ повернення'
  | 'Прихідна накладна';

type PrintDocument = {
  id: string;
  kind: DocumentKind;
  number: string;
  entityType: 'service_order' | 'sale' | 'purchase_order' | 'goods_receipt' | 'return' | 'contract' | 'act';
  entityId: string;
  clientOrSupplier: string;
  createdAt: string;
  createdBy: string;
  version: number;
  status: 'Чернетка' | 'Створено' | 'Роздруковано' | 'Підписано' | 'Скасовано' | 'Готово' | 'Очікує оплати' | 'Оплачено' | 'Готово до видачі' | 'PDF збережено';
  pdfPath: string;
  amountNet?: number;
  amountVat?: number;
  amountGross?: number;
  paymentId?: string;
  orderId?: string;
  invoiceId?: string;
  actId?: string;
  taxInvoiceId?: string;
  waybillId?: string;
  snapshotData?: string;
  orderVersionNo?: number;
  printedAt?: string;
  printedBy?: string;
  signedAt?: string;
  signedBy?: string;
};

type CompanyRequisites = {
  name: string;
  edrpou: string;
  iban: string;
  bank: string;
  mfo: string;
  address: string;
  phone: string;
  ipn: string;
  vatCertificate: string;
  taxStatus: string;
  signatory: string;
};

type CompanySettings = {
  companyName: string;
  mainEmail: string;
  techEmail: string;
  deprecatedEmail: string;
  edrpou: string;
  iban: string;
  bank: string;
  mfo: string;
  address: string;
  phone: string;
  ipn: string;
  vatCertificate: string;
  vatEnabled: boolean;
};

type ClientRequisites = {
  name: string;
  edrpou: string;
  address: string;
  contract: string;
  contact: string;
};

type ReleaseActStatus = 'Не створено' | 'Роздруковано' | 'Підписано';
type ActionType = 'Оплатити' | 'Підтвердити оплату' | 'Роздрукувати акт' | 'Підписати акт' | 'Видати' | 'Перейти до ремонту';

type NotificationChannel = 'SMS' | 'Telegram' | 'Viber' | 'Email';
type NotificationStatus = 'Створено' | 'Скопійовано' | 'Відправлено' | 'Помилка' | 'Вимкнено';
type NotificationEvent =
  | 'Прийом пристрою'
  | 'Діагностика розпочата'
  | 'Діагностика завершена'
  | 'Потрібне погодження додаткових робіт'
  | 'Погодження прийнято'
  | 'Погодження відхилено'
  | 'Очікування запчастини'
  | 'Ремонт розпочато'
  | 'Ремонт завершено'
  | 'Готово до видачі'
  | 'Очікує оплату'
  | 'Нагадування'
  | 'Видача'
  | 'Замовлення затримується';

type ClientNotification = {
  id: string;
  orderId: string;
  client: string;
  phone: string;
  channel: NotificationChannel;
  event: NotificationEvent;
  message: string;
  status: NotificationStatus;
  createdAt: string;
  createdBy?: string;
  templateName?: string;
  sentAt?: string;
  error?: string;
};

type ClientNotificationTemplate = {
  event: NotificationEvent;
  channel: NotificationChannel;
  enabled: boolean;
  text: string;
};

type SuggestedDocumentAction = {
  orderId: string;
  kind: 'Рахунок на оплату' | 'Акт надання послуг' | 'Наряд на ремонт' | 'Гарантійний талон';
  title: string;
  text: string;
  autoOpen?: boolean;
};

type ServiceOrderDocumentKind = 'Рахунок на оплату' | 'Акт надання послуг' | 'Видаткова накладна' | 'Наряд на ремонт' | 'Гарантійний талон';

type DashboardFocusTarget = 'approval' | 'payment' | 'issue';

type RepairApprovalStatus = 'Не відправлено' | 'Очікує відповідь' | 'Підтверджено' | 'Відмова' | 'Немає відповіді';
type RepairApproval = {
  id: string;
  orderId: string;
  token: string;
  client: string;
  phone: string;
  device: string;
  issue: string;
  worksSnapshot: Array<{ name: string; price: number; qty?: number; serviceType?: string }>;
  partsSnapshot: Array<{ name: string; qty: number; price: number }>;
  totalAmount: number;
  promisedDate: string;
  status: RepairApprovalStatus;
  sentAt: string;
  sentChannel: NotificationChannel;
  messageText: string;
  extraWorkDescription?: string;
  extraWorkAmount?: number;
  managerComment?: string;
  responseAt?: string;
  responseIp?: string;
  responseComment?: string;
  clientUrl: string;
};

type PayrollRuleType = 'Фіксована ставка' | 'За штуку' | 'Процент від роботи' | 'Ставка + процент';
type PayrollRule = {
  employee: string;
  role: Role;
  type: PayrollRuleType;
  monthlyRate?: number;
  perUnitRate?: number;
  percent?: number;
  bonus?: number;
  penaltyPerReturn?: number;
};

type Role = 'Руководитель' | 'Адміністратор' | 'Менеджер' | 'Інженер' | 'Комірник' | 'Бухгалтер' | 'Закупник';
type InternalMessageStatus = 'Нове' | 'Прочитано' | 'Виконано';
type InternalMessageImportance = 'Звичайна' | 'Важлива' | 'Критична';
type InternalMessage = {
  id: string;
  toUser: string;
  toRole: Role;
  orderId?: string;
  text: string;
  dueAt: string;
  importance: InternalMessageImportance;
  status: InternalMessageStatus;
  createdBy: string;
  createdAt: string;
};
type Permission =
  | 'clients.view'
  | 'clients.create'
  | 'clients.edit'
  | 'clients.archive'
  | 'repairs.view'
  | 'repairs.create'
  | 'repairs.edit'
  | 'repairs.status'
  | 'repairs.close'
  | 'repairs.cancel'
  | 'sales.view'
  | 'sales.create'
  | 'sales.edit'
  | 'sales.post'
  | 'sales.cancel'
  | 'sales.return'
  | 'stock.view'
  | 'stock.receive'
  | 'stock.reserve'
  | 'stock.unreserve'
  | 'stock.writeoff'
  | 'stock.inventory'
  | 'stock.adjust'
  | 'purchases.view'
  | 'purchases.create'
  | 'purchases.edit'
  | 'purchases.send'
  | 'purchases.status'
  | 'purchases.close'
  | 'finance.cash'
  | 'finance.payments'
  | 'finance.reports'
  | 'finance.refund'
  | 'finance.adjust'
  | 'reports.view'
  | 'reports.export'
  | 'reports.profit'
  | 'reports.staff'
  | 'settings.roles'
  | 'settings.users'
  | 'settings.templates'
  | 'settings.directories'
  | 'settings.system'
  | 'cost.view'
  | 'profit.view'
  | 'discount.approve.high'
  | 'documents.delete.posted'
  | 'documents.edit.closed'
  | 'payments.backdate.edit';

type AuthMode = 'phone_code' | 'email_password';
type UserPermissions = {
  canAccessWarehouse: boolean;
  canAccessFinance: boolean;
  canAccessEmployees: boolean;
  canAccessSettings: boolean;
  canAccessReports: boolean;
};

type User = {
  id: string;
  name: string;
  role: Role;
  twoFactor: boolean;
  login: string;
  phone?: string;
  email?: string;
  authMode: AuthMode;
  password?: string;
  permissions: UserPermissions;
  session: 'Активна' | 'Примусовий вихід' | 'Заблокована';
};

type LoginChallenge = {
  userId: string;
  code: string;
  phone: string;
  expiresAt: number;
  attempts: number;
};

type ActionLog = {
  id: string;
  date: string;
  user: string;
  role: Role;
  action: string;
  entity: string;
  comment: string;
};

type BackupPayload = {
  version: 1;
  users: User[];
  clients: ClientRecord[];
  products: Product[];
  orders: ServiceOrder[];
  payments: SimpleOrderPaymentRecord[];
  contracts: ContractRecord[];
  contractActs: ContractActRecord[];
  bankImportItems: BankImportItem[];
  requirements: PartRequirement[];
  purchases: PurchaseOrder[];
  receipts: GoodsReceipt[];
  movements: StockMovement[];
  taxInvoices: TaxInvoice[];
  actionLogs: ActionLog[];
  cashShift: CashShift;
};

type BackupSnapshot = {
  id: string;
  createdAt: string;
  createdAtTs: number;
  createdBy: string;
  source: 'auto' | 'manual' | 'import';
  label: string;
  payload: BackupPayload;
};

type ServiceProblemLevel = 'Критично' | 'Важливо' | 'Інформація';
type ServiceProblemType =
  | 'Заказ готовий, але не виданий'
  | 'Акт є, оплата відсутня'
  | 'Запчастини списані, але немає видаткової накладної'
  | 'Немає акта виконаних робіт'
  | 'Немає підстави для податкової накладної'
  | 'Заказ без руху'
  | 'Акт не повернений клієнтом'
  | 'Є заборгованість по заказу';

type ServiceProblem = {
  id: string;
  orderId: string;
  client: string;
  type: ServiceProblemType;
  description: string;
  days: number;
  responsible: string;
  engineer: string;
  manager: string;
  level: ServiceProblemLevel;
  target: 'order' | 'payment' | 'act' | 'document';
};

type AppToast = {
  id: string;
  message: string;
  tone: 'success' | 'error' | 'warning' | 'info';
};

const navItems: Array<{ id: Page; label: string; icon: React.ReactNode }> = [
  { id: 'dashboard', label: 'Головна', icon: <LayoutDashboard size={18} /> },
  { id: 'inbox', label: 'Inbox', icon: <Bell size={18} /> },
  { id: 'employee-control', label: 'Контроль співробітників', icon: <Users size={18} /> },
  { id: 'contracts', label: 'Договори', icon: <ClipboardList size={18} /> },
  { id: 'bank-import', label: 'Імпорт виписки', icon: <Banknote size={18} /> },
  { id: 'orders', label: 'Замовлення ремонту', icon: <ClipboardList size={18} /> },
  { id: 'my-orders', label: 'Мої замовлення', icon: <Wrench size={18} /> },
  { id: 'sales', label: 'Продажі', icon: <ShoppingCart size={18} /> },
  { id: 'clients', label: 'Клієнти', icon: <Users size={18} /> },
  { id: 'parts', label: 'Склад', icon: <Archive size={18} /> },
  { id: 'purchases', label: 'Закупівлі', icon: <ShoppingCart size={18} /> },
  { id: 'cash', label: 'Оплати / Каса', icon: <Banknote size={18} /> },
  { id: 'returns', label: 'Повернення', icon: <PackageCheck size={18} /> },
  { id: 'documents', label: 'Документи', icon: <ClipboardList size={18} /> },
  { id: 'tax-invoices', label: 'Податкові накладні', icon: <ClipboardList size={18} /> },
  { id: 'bas-exchange', label: 'Обмін з BAS', icon: <Archive size={18} /> },
  { id: 'problem-clients', label: 'Проблеми', icon: <Bell size={18} /> },
  { id: 'order-units', label: 'Одиниці заказів', icon: <PackageCheck size={18} /> },
  { id: 'storage', label: 'Полки / комірки', icon: <Archive size={18} /> },
  { id: 'movements', label: 'Рух складу', icon: <History size={18} /> },
  { id: 'finance', label: 'Фінанси', icon: <Banknote size={18} /> },
  { id: 'reports', label: 'Звіти', icon: <History size={18} /> },
  { id: 'payroll', label: 'Зарплата', icon: <Banknote size={18} /> },
  { id: 'acceptance', label: 'Чек-лист приймання', icon: <CheckCircle2 size={18} /> },
  { id: 'team', label: 'Співробітники', icon: <ShieldCheck size={18} /> },
  { id: 'settings', label: 'Налаштування', icon: <Settings size={18} /> },
];

const today = '21.04.2026';

const initialContracts: ContractRecord[] = [
  {
    id: 'DOG-1001',
    client: 'ТОВ "Альфа-Сервіс"',
    amount: 50000,
    startDate: '01.04.2026',
    endDate: '31.12.2026',
    status: 'Активний',
    createdAt: '01.04.2026',
  },
];

const companyRequisites: CompanyRequisites = {
  name: 'НВКПП «СПЕКТР-АС»',
  edrpou: '22590485',
  iban: 'UA123052990000026000000000001',
  bank: 'АТ КБ "ПриватБанк"',
  mfo: '305299',
  address: 'Україна, м. Київ, сервісний центр НВКПП «СПЕКТР-АС»',
  phone: '+380 44 000 00 00',
  ipn: '225904826000',
  vatCertificate: 'Свідоцтво платника ПДВ № 225904826000',
  taxStatus: 'Платник ПДВ, ІПН 225904826000',
  signatory: 'Директор',
};

const companySettings: CompanySettings = {
  companyName: 'СПЕКТР-АС',
  mainEmail: 'info.spektr.ac@ukr.net',
  techEmail: 'foppvg@ukr.net',
  deprecatedEmail: 'info@spektr.sumy.ua',
  edrpou: companyRequisites.edrpou,
  iban: companyRequisites.iban,
  bank: companyRequisites.bank,
  mfo: companyRequisites.mfo,
  address: companyRequisites.address,
  phone: companyRequisites.phone,
  ipn: companyRequisites.ipn,
  vatCertificate: companyRequisites.vatCertificate,
  vatEnabled: true,
};

const clientRequisites: Record<string, ClientRequisites> = {
  'ТОВ "Альфа-Сервіс"': { name: 'ТОВ "Альфа-Сервіс"', edrpou: '41234567', address: 'м. Київ, вул. Ділова, 15', contract: 'Договір № 12/04 від 12.04.2026', contact: '+380 44 301 88 20' },
  'Олена Коваленко': { name: 'Олена Коваленко', edrpou: 'фізична особа', address: 'м. Київ', contract: 'Без договору, згідно рахунку', contact: '+380 67 441 28 10' },
  'Ігор Мельник': { name: 'Ігор Мельник', edrpou: 'фізична особа', address: 'м. Київ', contract: 'Без договору, згідно рахунку', contact: '+380 50 617 73 04' },
};

const initialProducts: Product[] = [
  {
    id: 'p1',
    name: 'SSD Kingston NV2 500GB',
    sku: 'SSD-KIN-500',
    barcode: '4820001115008',
    extraBarcodes: [],
    category: 'Накопичувачі',
    brand: 'Kingston',
    model: 'NV2 500GB',
    unit: 'шт',
    cost: 1180,
    price: 1650,
    stock: 8,
    reserved: 0,
    withEngineer: 0,
    installed: 0,
    min: 3,
    storageLocation: 'A-01-01',
    batches: [
      { id: 'BAT-p1-1', productId: 'p1', qtyTotal: 3, qtyAvailable: 3, purchasePrice: 1120, purchaseDate: '10.02.2026', source: 'Prom', supplier: 'Prom', comment: 'Стара партія для FIFO' },
      { id: 'BAT-p1-2', productId: 'p1', qtyTotal: 5, qtyAvailable: 5, purchasePrice: 1180, purchaseDate: '05.04.2026', source: 'Постачальник Kingston', supplier: 'Постачальник Kingston' },
    ],
  },
  {
    id: 'p2',
    name: 'Матриця 15.6 IPS FHD',
    sku: 'LCD-156-FHD',
    barcode: '4820001561080',
    extraBarcodes: [],
    category: 'Екрани',
    brand: '',
    model: '15.6 IPS FHD',
    unit: 'шт',
    cost: 2550,
    price: 3200,
    stock: 0,
    reserved: 0,
    withEngineer: 0,
    installed: 0,
    min: 2,
    storageLocation: 'B-02-03',
    batches: [],
  },
  {
    id: 'p3',
    name: 'Ролик подачі Canon',
    sku: 'CAN-RL-02',
    barcode: '4820000020260',
    extraBarcodes: [],
    category: 'Оргтехніка',
    brand: 'Canon',
    model: 'RL-02',
    unit: 'шт',
    cost: 145,
    price: 260,
    stock: 11,
    reserved: 0,
    withEngineer: 0,
    installed: 0,
    min: 4,
    storageLocation: 'C-01-05',
    batches: [
      { id: 'BAT-p3-1', productId: 'p3', qtyTotal: 1, qtyAvailable: 1, purchasePrice: 100, purchaseDate: '01.01.2026', source: 'Вручну', supplier: 'Вручну', comment: 'Стара одиниця для тесту FIFO' },
      { id: 'BAT-p3-2', productId: 'p3', qtyTotal: 10, qtyAvailable: 10, purchasePrice: 120, purchaseDate: '01.02.2026', source: 'Prom', supplier: 'Prom', comment: 'Нова закупка' },
    ],
  },
  {
    id: 'p4',
    name: 'Акумулятор iPhone 13',
    sku: 'APL-BAT-13',
    barcode: '4820000131450',
    extraBarcodes: [],
    category: 'Мобільні',
    brand: 'Apple',
    model: 'iPhone 13',
    unit: 'шт',
    cost: 980,
    price: 1450,
    stock: 5,
    reserved: 0,
    withEngineer: 0,
    installed: 0,
    min: 2,
    storageLocation: 'M-03-02',
    batches: [
      { id: 'BAT-p4-1', productId: 'p4', qtyTotal: 2, qtyAvailable: 2, purchasePrice: 930, purchaseDate: '18.03.2026', source: 'Постачальник Apple Parts', supplier: 'Постачальник Apple Parts' },
      { id: 'BAT-p4-2', productId: 'p4', qtyTotal: 3, qtyAvailable: 3, purchasePrice: 980, purchaseDate: '15.04.2026', source: 'Prom', supplier: 'Prom' },
    ],
  },
];

const initialOrders: ServiceOrder[] = [
  {
    id: 'ЗН-1048',
    groupId: 'GR-20260419-1',
    qrUrl: '/orders/ЗН-1048',
    client: 'Олена Коваленко',
    phone: '+380 67 441 28 10',
    device: 'Ноутбук Lenovo ThinkPad E15',
    serial: 'PF4L9A21',
    issue: 'Не вмикається після перепаду напруги',
    engineer: 'Андрій Савчук',
    engineerAcceptedAt: '19.04.2026 09:35',
    manager: 'Вікторія Данилюк',
    legalEntity: false,
    status: 'На діагностиці',
    statusChangedAt: '19.04.2026 09:30',
    intakeDate: '19.04.2026 09:00',
    promisedDate: '22.04.2026 16:00',
    diagnosisResult: 'Перевіряється плата живлення та ланцюг заряджання.',
    deliveryAmount: 0,
    works: [{ name: 'Діагностика плати живлення', serviceType: 'Ремонт електроніки', price: 500, qty: 1, engineer: 'Андрій Савчук', accrualType: 'percent_of_work_amount', payrollPercent: 40 }],
    parts: [],
    payments: [],
    locationCode: 'A-3-1-1',
    locationStatus: 'У інженера',
    takenInRepairAt: '19.04.2026 09:35',
    takenInRepairBy: 'Андрій Савчук',
    takenFromLocation: 'A-3-1-1',
    urgent: true,
    clientNotified: false,
    lastNotificationType: undefined,
    lastNotificationAt: undefined,
    notificationHistory: [],
    pendingExtraApproval: false,
    statusHistory: [
      { id: 'H-1048-1', newStatus: 'Прийнято', changedBy: 'Вікторія Данилюк', changedAt: '19.04.2026 09:00', comment: 'Пристрій прийнято від клієнта.' },
      { id: 'H-1048-2', oldStatus: 'Прийнято', newStatus: 'На діагностиці', changedBy: 'Андрій Савчук', changedAt: '19.04.2026 09:30', comment: 'Передано інженеру.' },
    ],
  },
  {
    id: 'ЗН-1047',
    groupId: 'GR-20260418-1',
    qrUrl: '/orders/ЗН-1047',
    client: 'ТОВ "Альфа-Сервіс"',
    phone: '+380 44 301 88 20',
    device: 'МФУ Canon i-SENSYS MF445',
    serial: 'CAN-MF445-8821',
    issue: 'Зминання паперу, помилка подачі',
    engineer: 'Марина Лисенко',
    contractId: 'DOG-1001',
    engineerAcceptedAt: '18.04.2026 14:20',
    manager: 'Вікторія Данилюк',
    legalEntity: true,
    vatStatus: 'Очікує ПН',
    actIssuedAt: '19.04.2026',
    debtSince: '19.04.2026',
    status: 'В ремонті',
    statusChangedAt: '18.04.2026 14:20',
    intakeDate: '18.04.2026 10:15',
    promisedDate: '21.04.2026 12:00',
    diagnosisResult: 'Знос ролика подачі, потрібна профілактика вузла.',
    approvalComment: 'Клієнт погодив роботи і суму ремонту.',
    deliveryAmount: 0,
    works: [
      { name: 'Профілактика вузла подачі', serviceType: 'Ремонт оргтехніки', price: 850, qty: 1, engineer: 'Марина Лисенко', accrualType: 'percent_of_work_amount', payrollPercent: 25 },
      { name: 'Заправка картриджа Canon', serviceType: 'Заправка картриджів', price: 120, qty: 10, engineer: 'Марина Лисенко', accrualType: 'fixed_per_unit', payrollFixedPerUnit: 30 },
      { name: 'Регенерація картриджа', serviceType: 'Регенерація картриджів', price: 220, qty: 2, engineer: 'Марина Лисенко', accrualType: 'fixed_per_unit', payrollFixedPerUnit: 50 },
    ],
    parts: [],
    payments: [{ id: 'PAY-1', date: today, amount: 500, method: 'Картка', type: 'Передплата', transactionNo: 'TERM-0001', acceptedBy: 'Вікторія Данилюк', orderId: 'ЗН-1047', comment: 'Передплата за ремонт' }],
    locationCode: 'B-2-1-3',
    locationStatus: 'У інженера',
    takenInRepairAt: '18.04.2026 14:20',
    takenInRepairBy: 'Марина Лисенко',
    takenFromLocation: 'B-2-1-3',
    urgent: false,
    clientNotified: true,
    lastNotificationType: 'Прийом пристрою',
    lastNotificationAt: '18.04.2026 10:15',
    notificationHistory: [],
    pendingExtraApproval: false,
    statusHistory: [
      { id: 'H-1047-1', newStatus: 'Прийнято', changedBy: 'Вікторія Данилюк', changedAt: '18.04.2026 10:15', comment: 'Пристрій прийнято.' },
      { id: 'H-1047-2', oldStatus: 'Прийнято', newStatus: 'На діагностиці', changedBy: 'Марина Лисенко', changedAt: '18.04.2026 11:00', comment: 'Початок діагностики.' },
      { id: 'H-1047-3', oldStatus: 'На діагностиці', newStatus: 'Погоджено', changedBy: 'Вікторія Данилюк', changedAt: '18.04.2026 13:50', comment: 'Клієнт погодив ремонт.' },
      { id: 'H-1047-4', oldStatus: 'Погоджено', newStatus: 'В ремонті', changedBy: 'Марина Лисенко', changedAt: '18.04.2026 14:20', comment: 'Інженер почав ремонт.' },
    ],
  },
  {
    id: 'ЗН-1050',
    groupId: 'GR-20260501-DEMO-VAT',
    qrUrl: '/orders/ЗН-1050',
    client: 'ТОВ "Демо ПДВ"',
    phone: '+380 44 555 10 50',
    device: 'Принтер HP LaserJet Pro M404',
    serial: 'VAT-DEMO-1050',
    issue: 'TEST ПН · до створення',
    engineer: 'Марина Лисенко',
    engineerAcceptedAt: '30.04.2026 10:40',
    manager: 'Вікторія Данилюк',
    legalEntity: true,
    vatStatus: 'Очікує ПН',
    actIssuedAt: '01.05.2026',
    status: 'Готовий до видачі',
    statusChangedAt: '01.05.2026 15:20',
    intakeDate: '30.04.2026 09:50',
    promisedDate: '02.05.2026 18:00',
    diagnosisResult: 'Профілактика тракту друку та очищення вузлів.',
    approvalComment: 'Демо-замовлення для перевірки ПН до створення.',
    deliveryAmount: 0,
    works: [
      { name: 'Сервісне обслуговування принтера', serviceType: 'Ремонт оргтехніки', price: 2400, qty: 1, engineer: 'Марина Лисенко', accrualType: 'percent_of_work_amount', payrollPercent: 25 },
    ],
    parts: [],
    payments: [
      {
        id: 'PAY-DEMO-1050',
        date: '01.05.2026 15:00',
        amount: 2400,
        method: 'Картка',
        type: 'Повна оплата',
        transactionNo: 'TERM-DEMO-1050',
        acceptedBy: 'Вікторія Данилюк',
        status: 'Проведено',
        orderId: 'ЗН-1050',
        comment: 'Демо-оплата для створення ПН',
      },
    ],
    locationCode: 'R-1-0-5-0',
    locationStatus: 'У комірці',
    takenInRepairAt: '30.04.2026 10:40',
    takenInRepairBy: 'Марина Лисенко',
    takenFromLocation: 'R-1-0-5-0',
    returnedToCellAt: '01.05.2026 15:20',
    returnedToCellBy: 'Марина Лисенко',
    urgent: false,
    clientNotified: true,
    lastNotificationType: 'Готово до видачі',
    lastNotificationAt: '01.05.2026 15:25',
    notificationHistory: [],
    pendingExtraApproval: false,
    engineerWorkCompletedAt: '01.05.2026 14:55',
    engineerWorkCompletedBy: 'Марина Лисенко',
    currentVersion: 1,
    statusHistory: [
      { id: 'H-1050-1', newStatus: 'Прийнято', changedBy: 'Вікторія Данилюк', changedAt: '30.04.2026 09:50', comment: 'Демо-замовлення прийнято.' },
      { id: 'H-1050-2', oldStatus: 'Прийнято', newStatus: 'На діагностиці', changedBy: 'Марина Лисенко', changedAt: '30.04.2026 10:40', comment: 'Інженер взяв у роботу.' },
      { id: 'H-1050-3', oldStatus: 'На діагностиці', newStatus: 'В ремонті', changedBy: 'Марина Лисенко', changedAt: '30.04.2026 11:10', comment: 'Розпочато ремонт.' },
      { id: 'H-1050-4', oldStatus: 'В ремонті', newStatus: 'Готовий до видачі', changedBy: 'Вікторія Данилюк', changedAt: '01.05.2026 15:20', comment: 'Оплата проведена, акт підписано, ПН ще не створено.' },
    ],
  },
  {
    id: 'ЗН-1051',
    groupId: 'GR-20260501-DEMO-VAT',
    qrUrl: '/orders/ЗН-1051',
    client: 'ТОВ "TEST ПН"',
    phone: '+380 44 555 10 51',
    device: 'МФУ Canon i-SENSYS MF443',
    serial: 'VAT-DEMO-1051',
    issue: 'TEST ПН · зареєстровано',
    engineer: 'Марина Лисенко',
    engineerAcceptedAt: '30.04.2026 12:10',
    manager: 'Вікторія Данилюк',
    legalEntity: true,
    vatStatus: 'ПН зареєстрована',
    actIssuedAt: '01.05.2026',
    status: 'Готовий до видачі',
    statusChangedAt: '01.05.2026 17:10',
    intakeDate: '30.04.2026 11:30',
    promisedDate: '02.05.2026 18:30',
    diagnosisResult: 'Повна профілактика та заміна витратних вузлів не потрібна.',
    approvalComment: 'Демо-замовлення для перевірки створеної та зареєстрованої ПН.',
    deliveryAmount: 0,
    works: [
      { name: 'Сервісне обслуговування МФУ', serviceType: 'Ремонт оргтехніки', price: 3180, qty: 1, engineer: 'Марина Лисенко', accrualType: 'percent_of_work_amount', payrollPercent: 25 },
    ],
    parts: [],
    payments: [
      {
        id: 'PAY-DEMO-1051',
        date: '01.05.2026 16:35',
        amount: 3180,
        method: 'Картка',
        type: 'Повна оплата',
        transactionNo: 'TERM-DEMO-1051',
        acceptedBy: 'Вікторія Данилюк',
        status: 'Проведено',
        orderId: 'ЗН-1051',
        comment: 'Демо-оплата для повного циклу ПН',
      },
    ],
    locationCode: 'R-1-0-5-1',
    locationStatus: 'У комірці',
    takenInRepairAt: '30.04.2026 12:10',
    takenInRepairBy: 'Марина Лисенко',
    takenFromLocation: 'R-1-0-5-1',
    returnedToCellAt: '01.05.2026 17:10',
    returnedToCellBy: 'Марина Лисенко',
    urgent: false,
    clientNotified: true,
    lastNotificationType: 'Готово до видачі',
    lastNotificationAt: '01.05.2026 17:12',
    notificationHistory: [],
    pendingExtraApproval: false,
    engineerWorkCompletedAt: '01.05.2026 16:20',
    engineerWorkCompletedBy: 'Марина Лисенко',
    currentVersion: 1,
    statusHistory: [
      { id: 'H-1051-1', newStatus: 'Прийнято', changedBy: 'Вікторія Данилюк', changedAt: '30.04.2026 11:30', comment: 'Демо-замовлення прийнято.' },
      { id: 'H-1051-2', oldStatus: 'Прийнято', newStatus: 'На діагностиці', changedBy: 'Марина Лисенко', changedAt: '30.04.2026 12:10', comment: 'Інженер взяв у роботу.' },
      { id: 'H-1051-3', oldStatus: 'На діагностиці', newStatus: 'В ремонті', changedBy: 'Марина Лисенко', changedAt: '30.04.2026 12:45', comment: 'Розпочато ремонт.' },
      { id: 'H-1051-4', oldStatus: 'В ремонті', newStatus: 'Готовий до видачі', changedBy: 'Вікторія Данилюк', changedAt: '01.05.2026 17:10', comment: 'Повний цикл ПН завершено.' },
    ],
  },
];

const initialSales: Sale[] = [
  {
    id: 'ПР-2001',
    client: 'Наталія Романюк',
    manager: 'Вікторія Данилюк',
    date: today,
    status: 'Чернетка',
    serviceOrderId: 'ЗН-1047',
    items: [],
    payments: [],
    returns: [],
  },
];

const initialDocuments: PrintDocument[] = [
  { id: 'DOC-1', kind: 'Квитанція прийому', number: 'КП-1048', entityType: 'service_order', entityId: 'ЗН-1048', clientOrSupplier: 'Олена Коваленко', createdAt: '19.04.2026 09:00', createdBy: 'Вікторія Данилюк', version: 1, status: 'PDF збережено', pdfPath: '/documents/service/ZN-1048/intake.pdf' },
  { id: 'DOC-2', kind: 'Заказ-наряд', number: 'ЗНД-1047', entityType: 'service_order', entityId: 'ЗН-1047', clientOrSupplier: 'ТОВ "Альфа-Сервіс"', createdAt: '18.04.2026 13:50', createdBy: 'Вікторія Данилюк', version: 1, status: 'PDF збережено', pdfPath: '/documents/service/ZN-1047/work-order.pdf' },
  {
    id: 'DOC-DEMO-1050-INV',
    kind: 'Рахунок на оплату',
    number: 'РХ-1050',
    entityType: 'service_order',
    entityId: 'ЗН-1050',
    clientOrSupplier: 'ТОВ "Демо ПДВ"',
    createdAt: '01.05.2026 14:40',
    createdBy: 'Вікторія Данилюк',
    version: 1,
    status: 'Роздруковано',
    pdfPath: '/documents/service/ZN-1050/invoice.pdf',
    amountNet: 2000,
    amountVat: 400,
    amountGross: 2400,
    orderId: 'ЗН-1050',
    invoiceId: 'DOC-DEMO-1050-INV',
    orderVersionNo: 1,
    printedAt: '01.05.2026 14:42',
    printedBy: 'Вікторія Данилюк',
    snapshotData: JSON.stringify({
      date: '01.05.2026',
      orderNumber: 'РХ-1050',
      clientName: 'ТОВ "Демо ПДВ"',
      phone: '+380 44 555 10 50',
      device: 'Принтер HP LaserJet Pro M404',
      serialNumber: 'VAT-DEMO-1050',
      problem: 'TEST ПН · до створення',
      total: '2400 грн',
      paid: '2400 грн',
      debt: '0 грн',
      managerName: 'Вікторія Данилюк',
      engineerName: 'Марина Лисенко',
      clientTaxId: '44556677',
      amountNet: '2000 грн',
      amountVat: '400 грн',
      amountGross: '2400 грн',
      amountWords: 'Дві тисячі чотириста гривень 00 копійок',
      vatWords: 'Чотириста гривень 00 копійок',
      companyName: 'ТОВ "Сервіс+"',
      companyEdrpou: '22590485',
      companyIban: 'UA123456789012345678901234567',
      companyBank: 'АТ КБ "ПриватБанк"',
      companyMfo: '305299',
      companyAddress: 'м. Київ, вул. Хрещатик, 1',
      companyPhone: '+380 44 555 12 34',
      companyIpn: '225904826000',
      companyVatCertificate: 'ПДВ № 100200300',
      companyVatStatus: 'Платник ПДВ',
      invoiceReason: 'Демо-рахунок для створення ПН',
      totalWithVATWords: 'Дві тисячі чотириста гривень 00 копійок',
      vatAmountWords: 'Чотириста гривень 00 копійок',
    }),
  },
  {
    id: 'DOC-DEMO-1050-ACT',
    kind: 'Акт надання послуг',
    number: 'АНП-1050',
    entityType: 'service_order',
    entityId: 'ЗН-1050',
    clientOrSupplier: 'ТОВ "Демо ПДВ"',
    createdAt: '01.05.2026 15:05',
    createdBy: 'Вікторія Данилюк',
    version: 1,
    status: 'Підписано',
    pdfPath: '/documents/service/ZN-1050/act.pdf',
    amountNet: 2000,
    amountVat: 400,
    amountGross: 2400,
    orderId: 'ЗН-1050',
    invoiceId: 'DOC-DEMO-1050-INV',
    actId: 'DOC-DEMO-1050-ACT',
    orderVersionNo: 1,
    printedAt: '01.05.2026 15:06',
    printedBy: 'Вікторія Данилюк',
    signedAt: '01.05.2026 15:15',
    signedBy: 'Вікторія Данилюк',
    snapshotData: JSON.stringify({
      date: '01.05.2026',
      orderNumber: 'АНП-1050',
      clientName: 'ТОВ "Демо ПДВ"',
      phone: '+380 44 555 10 50',
      device: 'Принтер HP LaserJet Pro M404',
      serialNumber: 'VAT-DEMO-1050',
      problem: 'TEST ПН · до створення',
      total: '2400 грн',
      paid: '2400 грн',
      debt: '0 грн',
      managerName: 'Вікторія Данилюк',
      engineerName: 'Марина Лисенко',
      clientTaxId: '44556677',
      amountNet: '2000 грн',
      amountVat: '400 грн',
      amountGross: '2400 грн',
      amountWords: 'Дві тисячі чотириста гривень 00 копійок',
      vatWords: 'Чотириста гривень 00 копійок',
      companyName: 'ТОВ "Сервіс+"',
      companyEdrpou: '22590485',
      companyIban: 'UA123456789012345678901234567',
      companyBank: 'АТ КБ "ПриватБанк"',
      companyMfo: '305299',
      companyAddress: 'м. Київ, вул. Хрещатик, 1',
      companyPhone: '+380 44 555 12 34',
      companyIpn: '225904826000',
      companyVatCertificate: 'ПДВ № 100200300',
      companyVatStatus: 'Платник ПДВ',
      linkedInvoiceNumber: 'РХ-1050',
      linkedInvoiceDate: '01.05.2026',
      totalWithVATWords: 'Дві тисячі чотириста гривень 00 копійок',
      vatAmountWords: 'Чотириста гривень 00 копійок',
      serviceItemsRows: '<tr><td class="center">1</td><td>Сервісне обслуговування принтера</td><td class="center">1</td><td class="right">2400 грн</td><td class="right">2400 грн</td></tr>',
      productItemsRows: '<tr><td class="center">1</td><td>Комплектуючі не використовувались</td><td class="center">0</td><td class="right">0 грн</td><td class="right">0 грн</td></tr>',
    }),
  },
  {
    id: 'DOC-DEMO-1051-INV',
    kind: 'Рахунок на оплату',
    number: 'РХ-1051',
    entityType: 'service_order',
    entityId: 'ЗН-1051',
    clientOrSupplier: 'ТОВ "TEST ПН"',
    createdAt: '01.05.2026 16:10',
    createdBy: 'Вікторія Данилюк',
    version: 1,
    status: 'Роздруковано',
    pdfPath: '/documents/service/ZN-1051/invoice.pdf',
    amountNet: 2650,
    amountVat: 530,
    amountGross: 3180,
    orderId: 'ЗН-1051',
    invoiceId: 'DOC-DEMO-1051-INV',
    taxInvoiceId: 'TIN-DEMO-1051',
    orderVersionNo: 1,
    printedAt: '01.05.2026 16:12',
    printedBy: 'Вікторія Данилюк',
    snapshotData: JSON.stringify({
      date: '01.05.2026',
      orderNumber: 'РХ-1051',
      clientName: 'ТОВ "TEST ПН"',
      phone: '+380 44 555 10 51',
      device: 'МФУ Canon i-SENSYS MF443',
      serialNumber: 'VAT-DEMO-1051',
      problem: 'TEST ПН · зареєстровано',
      total: '3180 грн',
      paid: '3180 грн',
      debt: '0 грн',
      managerName: 'Вікторія Данилюк',
      engineerName: 'Марина Лисенко',
      clientTaxId: '44556678',
      amountNet: '2650 грн',
      amountVat: '530 грн',
      amountGross: '3180 грн',
      amountWords: 'Три тисячі сто вісімдесят гривень 00 копійок',
      vatWords: 'П’ятсот тридцять гривень 00 копійок',
      companyName: 'ТОВ "Сервіс+"',
      companyEdrpou: '22590485',
      companyIban: 'UA123456789012345678901234567',
      companyBank: 'АТ КБ "ПриватБанк"',
      companyMfo: '305299',
      companyAddress: 'м. Київ, вул. Хрещатик, 1',
      companyPhone: '+380 44 555 12 34',
      companyIpn: '225904826000',
      companyVatCertificate: 'ПДВ № 100200300',
      companyVatStatus: 'Платник ПДВ',
      invoiceReason: 'Демо-рахунок для повного циклу ПН',
      totalWithVATWords: 'Три тисячі сто вісімдесят гривень 00 копійок',
      vatAmountWords: 'П’ятсот тридцять гривень 00 копійок',
    }),
  },
  {
    id: 'DOC-DEMO-1051-ACT',
    kind: 'Акт надання послуг',
    number: 'АНП-1051',
    entityType: 'service_order',
    entityId: 'ЗН-1051',
    clientOrSupplier: 'ТОВ "TEST ПН"',
    createdAt: '01.05.2026 16:40',
    createdBy: 'Вікторія Данилюк',
    version: 1,
    status: 'Підписано',
    pdfPath: '/documents/service/ZN-1051/act.pdf',
    amountNet: 2650,
    amountVat: 530,
    amountGross: 3180,
    orderId: 'ЗН-1051',
    invoiceId: 'DOC-DEMO-1051-INV',
    actId: 'DOC-DEMO-1051-ACT',
    taxInvoiceId: 'TIN-DEMO-1051',
    orderVersionNo: 1,
    printedAt: '01.05.2026 16:42',
    printedBy: 'Вікторія Данилюк',
    signedAt: '01.05.2026 16:50',
    signedBy: 'Вікторія Данилюк',
    snapshotData: JSON.stringify({
      date: '01.05.2026',
      orderNumber: 'АНП-1051',
      clientName: 'ТОВ "TEST ПН"',
      phone: '+380 44 555 10 51',
      device: 'МФУ Canon i-SENSYS MF443',
      serialNumber: 'VAT-DEMO-1051',
      problem: 'TEST ПН · зареєстровано',
      total: '3180 грн',
      paid: '3180 грн',
      debt: '0 грн',
      managerName: 'Вікторія Данилюк',
      engineerName: 'Марина Лисенко',
      clientTaxId: '44556678',
      amountNet: '2650 грн',
      amountVat: '530 грн',
      amountGross: '3180 грн',
      amountWords: 'Три тисячі сто вісімдесят гривень 00 копійок',
      vatWords: 'П’ятсот тридцять гривень 00 копійок',
      companyName: 'ТОВ "Сервіс+"',
      companyEdrpou: '22590485',
      companyIban: 'UA123456789012345678901234567',
      companyBank: 'АТ КБ "ПриватБанк"',
      companyMfo: '305299',
      companyAddress: 'м. Київ, вул. Хрещатик, 1',
      companyPhone: '+380 44 555 12 34',
      companyIpn: '225904826000',
      companyVatCertificate: 'ПДВ № 100200300',
      companyVatStatus: 'Платник ПДВ',
      linkedInvoiceNumber: 'РХ-1051',
      linkedInvoiceDate: '01.05.2026',
      totalWithVATWords: 'Три тисячі сто вісімдесят гривень 00 копійок',
      vatAmountWords: 'П’ятсот тридцять гривень 00 копійок',
      serviceItemsRows: '<tr><td class="center">1</td><td>Сервісне обслуговування МФУ</td><td class="center">1</td><td class="right">3180 грн</td><td class="right">3180 грн</td></tr>',
      productItemsRows: '<tr><td class="center">1</td><td>Комплектуючі не використовувались</td><td class="center">0</td><td class="right">0 грн</td><td class="right">0 грн</td></tr>',
    }),
  },
];

const documentTemplates: Array<{ kind: DocumentKind; description: string; lockedRule: string }> = [
  { kind: 'Квитанція прийому', description: 'Фіксує стан, комплектність, дефект зі слів клієнта і підписи.', lockedRule: 'Після підпису не редагується.' },
  { kind: 'Акт діагностики', description: 'Містить результат діагностики, рекомендації, роботи, запчастини, суму і строк.', lockedRule: 'Версія зберігається в історії.' },
  { kind: 'Заказ-наряд', description: 'Основний документ ремонту з роботами, запчастинами, статусом і підсумком.', lockedRule: 'Суми не змінюються заднім числом.' },
  { kind: 'Рахунок на оплату', description: 'Офіційний рахунок з реквізитами сторін, договором, позиціями, сумами без ПДВ / ПДВ / з ПДВ і сумою прописью.', lockedRule: 'Розрахунки виконуються системою, шаблон редагується окремо.' },
  { kind: 'Акт надання послуг', description: 'Офіційний акт із розділенням виконаних робіт / послуг і використаних комплектуючих, реквізитами та підписами сторін.', lockedRule: 'Після підпису створюється нова версія, стара не змінюється.' },
  { kind: 'Акт виконаних робіт', description: 'Підтверджує виконані роботи, встановлені деталі, гарантію і підпис клієнта.', lockedRule: 'Після видачі блокується.' },
  { kind: 'Акт видачі', description: 'Фіксує факт видачі пристрою клієнту та відсутність претензій.', lockedRule: 'Видалення тільки з правом адміністратора.' },
  { kind: 'Видаткова накладна', description: 'Створюється тільки для товарів і комплектуючих зі складу, не для чистих послуг.', lockedRule: 'Не створюється без товарних позицій.' },
  { kind: 'Наряд на ремонт', description: 'Ремонтний наряд із пристроєм, серійним номером, несправністю, роботами, запчастинами і підписами.', lockedRule: 'Створюється з картки ремонту.' },
  { kind: 'Швидкий наряд', description: 'Короткий наряд для картриджів: клієнт, картридж, вид роботи, результат і підписи.', lockedRule: 'Використовується для потоку картриджів.' },
  { kind: 'Акт технічного стану', description: 'Фіксує дефекти, комісію і висновок "не підлягає ремонту".', lockedRule: 'Обовʼязковий для статусу не підлягає ремонту.' },
  { kind: 'Гарантійний талон', description: 'Містить пристрій, серійний номер, дату, термін гарантії та умови.', lockedRule: 'Друкується після завершення ремонту.' },
  { kind: 'Товарний чек', description: 'Документ продажу з товарами, цінами, знижками, оплатою і підсумком.', lockedRule: 'Після оплати створюється PDF.' },
  { kind: 'Накладна продажу', description: 'Печатна накладна для відвантаження товару клієнту.', lockedRule: 'Після видачі блокується.' },
  { kind: 'Касовий чек', description: 'Внутрішній або інтегрований касовий документ по оплаті.', lockedRule: 'Звʼязується з платежем.' },
  { kind: 'Документ повернення', description: 'Повернення товару або сервісної позиції з причиною і сумою.', lockedRule: 'Потрібне посилання на продаж або ремонт.' },
  { kind: 'Прихідна накладна', description: 'Складський документ приходу від постачальника.', lockedRule: 'Створює рух складу після проведення.' },
];

const documentPrefix: Record<DocumentKind, string> = {
  'Квитанція прийому': 'КП',
  'Акт діагностики': 'АД',
  'Заказ-наряд': 'ЗНД',
  'Рахунок на оплату': 'РХ',
  'Акт надання послуг': 'АНП',
  'Акт виконаних робіт': 'АВР',
  'Акт видачі': 'АВ',
  'Видаткова накладна': 'ВН',
  'Наряд на ремонт': 'НР',
  'Швидкий наряд': 'ШН',
  'Акт технічного стану': 'АТС',
  'Гарантійний талон': 'ГТ',
  'Товарний чек': 'ТЧ',
  'Накладна продажу': 'НП',
  'Касовий чек': 'КЧ',
  'Документ повернення': 'ПВ',
  'Прихідна накладна': 'ПН',
};

const clientNotificationTemplates: ClientNotificationTemplate[] = [
  { event: 'Прийом пристрою', channel: 'SMS', enabled: true, text: 'Ваше замовлення №{order} прийнято. Пристрій: {device}. Очікуйте діагностику.' },
  { event: 'Діагностика розпочата', channel: 'SMS', enabled: true, text: 'Замовлення №{order}. Діагностику пристрою {device} розпочато.' },
  { event: 'Діагностика завершена', channel: 'SMS', enabled: true, text: 'Замовлення №{order}. Виявлено проблему: {problem}. Вартість ремонту: {amount}. Відповідь: 1 — згоден, 2 — відмовитись.' },
  { event: 'Потрібне погодження додаткових робіт', channel: 'SMS', enabled: true, text: 'Додаткові роботи по замовленню №{order}: {extraWorkDescription}. Сума: {extraWorkAmount}. Відповідь: 1 — згоден, 2 — відмова.' },
  { event: 'Погодження прийнято', channel: 'Viber', enabled: true, text: 'Замовлення №{order}. Додаткові роботи погоджено. Загальна сума: {amount}.' },
  { event: 'Погодження відхилено', channel: 'Viber', enabled: true, text: 'Замовлення №{order}. Відмову від додаткових робіт зафіксовано. Телефон сервісу: {servicePhone}.' },
  { event: 'Очікування запчастини', channel: 'SMS', enabled: true, text: 'Замовлення №{order}. Очікуємо запчастину. Орієнтовний строк: {date}.' },
  { event: 'Ремонт розпочато', channel: 'SMS', enabled: true, text: 'Замовлення №{order}. Ваш заказ уже в роботі: {device}.' },
  { event: 'Ремонт завершено', channel: 'SMS', enabled: true, text: 'Замовлення №{order}. Ремонт завершено. Сума: {amount}.' },
  { event: 'Готово до видачі', channel: 'SMS', enabled: true, text: 'Замовлення №{order} готове. Сума до оплати: {amount}. Забрати: {serviceAddress}. Працюємо до {serviceHours}.' },
  { event: 'Очікує оплату', channel: 'SMS', enabled: true, text: 'Замовлення №{order} очікує оплату. До сплати: {amount}.' },
  { event: 'Нагадування', channel: 'SMS', enabled: true, text: 'Замовлення №{order} готове та чекає на вас. Через 5 днів нараховується зберігання.' },
  { event: 'Видача', channel: 'SMS', enabled: true, text: 'Замовлення №{order} видано. Дякуємо за звернення.' },
  { event: 'Замовлення затримується', channel: 'SMS', enabled: true, text: 'Замовлення №{order} зберігається понад {days} днів. Просимо забрати найближчим часом.' },
  { event: 'Готово до видачі', channel: 'Email', enabled: false, text: 'Ваше замовлення {order} готове до видачі. Сума: {amount}.' },
];

const initialClientNotifications: ClientNotification[] = [
  {
    id: 'MSG-1',
    orderId: 'ЗН-1047',
    client: 'ТОВ "Альфа-Сервіс"',
    phone: '+380 44 301 88 20',
    channel: 'SMS',
    event: 'Прийом пристрою',
    message: 'ТОВ "Альфа-Сервіс", ваш пристрій прийнято. Номер замовлення: ЗН-1047.',
    status: 'Відправлено',
    createdAt: '18.04.2026 10:15',
    sentAt: '18.04.2026 10:15',
  },
];

const initialRepairApprovals: RepairApproval[] = [
  {
    id: 'APR-1',
    orderId: 'ЗН-1047',
    token: 'appr-1047-demo',
    client: 'ТОВ "Альфа-Сервіс"',
    phone: '+380 44 301 88 20',
    device: 'МФУ Canon i-SENSYS MF445',
    issue: 'Зажовує папір',
    worksSnapshot: [{ name: 'Профілактика вузла подачі', price: 850, qty: 1, serviceType: 'Ремонт оргтехніки' }],
    partsSnapshot: [],
    totalAmount: 850,
    promisedDate: '21.04.2026 12:00',
    status: 'Підтверджено',
    sentAt: '18.04.2026 13:50',
    sentChannel: 'Telegram',
    messageText: 'Вартість ремонту ЗН-1047: 850 грн. Підтвердіть за посиланням.',
    responseAt: '18.04.2026 14:05',
    responseIp: '93.77.120.14',
    responseComment: 'Клієнт підтвердив ремонт через посилання.',
    clientUrl: '/approval/appr-1047-demo',
  },
];

const initialTaxInvoices: TaxInvoice[] = [
  {
    id: 'TIN-1047',
    number: 'ПН-1047',
    date: '19.04.2026',
    orderId: 'ЗН-1047',
    actId: 'DOC-1047-ACT',
    eventType: 'act',
    client: 'ТОВ "Альфа-Сервіс"',
    eventDate: '19.04.2026',
    amount: 2490,
    vatAmount: Math.round(vatFromGross(2490).vat),
    status: 'Створено',
    snapshot: JSON.stringify({
      companyName: 'ТОВ "Сервіс+"',
      companyEdrpou: '22590485',
      companyIpn: '225904826000',
      companyIban: 'UA123456789012345678901234567',
      companyBank: 'АТ КБ "ПриватБанк"',
      companyMfo: '305299',
      companyAddress: 'м. Київ, вул. Хрещатик, 1',
      companyPhone: '+380 44 555 12 34',
      companyVatStatus: 'Платник ПДВ',
      clientName: 'ТОВ "Альфа-Сервіс"',
      clientPhone: '+380 44 301 88 20',
      clientTaxId: '22590485',
      orderId: 'ЗН-1047',
      orderNumber: 'ЗН-1047',
      device: 'МФУ Canon i-SENSYS MF445',
      managerName: 'Ірина Бондар',
      documentNumber: 'ПН-1047',
      documentDate: '19.04.2026',
      eventDate: '19.04.2026',
      eventType: 'act',
      sumWithoutVAT: 2075,
      vatAmount: 415,
      sumWithVAT: 2490,
      totalWithVATWords: 'Дві тисячі чотириста дев’яносто гривень 00 копійок',
      vatAmountWords: 'Чотириста п’ятнадцять гривень 00 копійок',
      lines: [],
      actId: 'DOC-1047-ACT',
      actNumber: 'ACT-1047',
      createdAt: '19.04.2026 16:00',
    } satisfies TaxInvoiceSnapshot),
    createdAt: '19.04.2026 16:00',
    responsible: 'Ірина Бондар',
  },
  {
    id: 'TIN-1039',
    number: 'ПН-1039',
    date: '10.04.2026',
    orderId: 'ЗН-1039',
    paymentId: 'PAY-1039',
    eventType: 'payment',
    client: 'ТОВ "Офіс-Друк"',
    eventDate: '10.04.2026',
    amount: 7200,
    vatAmount: Math.round(vatFromGross(7200).vat),
    status: 'Помилка',
    snapshot: JSON.stringify({
      companyName: 'ТОВ "Сервіс+"',
      companyEdrpou: '22590485',
      companyIpn: '225904826000',
      companyIban: 'UA123456789012345678901234567',
      companyBank: 'АТ КБ "ПриватБанк"',
      companyMfo: '305299',
      companyAddress: 'м. Київ, вул. Хрещатик, 1',
      companyPhone: '+380 44 555 12 34',
      companyVatStatus: 'Платник ПДВ',
      clientName: 'ТОВ "Офіс-Друк"',
      clientPhone: '+380 44 777 12 11',
      clientTaxId: '30111223',
      orderId: 'ЗН-1039',
      orderNumber: 'ЗН-1039',
      device: 'Принтер HP LaserJet',
      managerName: 'Ірина Бондар',
      documentNumber: 'ПН-1039',
      documentDate: '10.04.2026',
      eventDate: '10.04.2026',
      eventType: 'payment',
      sumWithoutVAT: 6000,
      vatAmount: 1200,
      sumWithVAT: 7200,
      totalWithVATWords: 'Сім тисяч двісті гривень 00 копійок',
      vatAmountWords: 'Одна тисяча двісті гривень 00 копійок',
      lines: [],
      paymentId: 'PAY-1039',
      paymentType: 'Безготівка',
      paymentStatus: 'Підтверджено',
      createdAt: '10.04.2026 12:30',
    } satisfies TaxInvoiceSnapshot),
    createdAt: '10.04.2026 12:30',
    registeredAt: undefined,
    responsible: 'Ірина Бондар',
  },
  {
    id: 'TIN-DEMO-1051',
    number: 'ПН-1051',
    date: '01.05.2026',
    orderId: 'ЗН-1051',
    invoiceId: 'DOC-DEMO-1051-INV',
    actId: 'DOC-DEMO-1051-ACT',
    paymentId: 'PAY-DEMO-1051',
    eventType: 'payment',
    client: 'ТОВ "TEST ПН"',
    eventDate: '01.05.2026',
    amount: 3180,
    vatAmount: 530,
    status: 'Зареєстровано',
    snapshot: JSON.stringify({
      companyName: 'ТОВ "Сервіс+"',
      companyEdrpou: '22590485',
      companyIpn: '225904826000',
      companyIban: 'UA123456789012345678901234567',
      companyBank: 'АТ КБ "ПриватБанк"',
      companyMfo: '305299',
      companyAddress: 'м. Київ, вул. Хрещатик, 1',
      companyPhone: '+380 44 555 12 34',
      companyVatStatus: 'Платник ПДВ',
      clientName: 'ТОВ "TEST ПН"',
      clientPhone: '+380 44 555 10 51',
      clientTaxId: '44556678',
      orderId: 'ЗН-1051',
      orderNumber: 'ЗН-1051',
      device: 'МФУ Canon i-SENSYS MF443',
      managerName: 'Вікторія Данилюк',
      documentNumber: 'ПН-1051',
      documentDate: '01.05.2026',
      eventDate: '01.05.2026',
      eventType: 'payment',
      sumWithoutVAT: 2650,
      vatAmount: 530,
      sumWithVAT: 3180,
      totalWithVATWords: 'Три тисячі сто вісімдесят гривень 00 копійок',
      vatAmountWords: 'П’ятсот тридцять гривень 00 копійок',
      lines: [
        { type: 'service', label: 'Сервісне обслуговування МФУ', qty: 1, price: 3180, total: 3180 },
      ],
      paymentId: 'PAY-DEMO-1051',
      paymentType: 'Картка',
      paymentStatus: 'Проведено',
      invoiceId: 'DOC-DEMO-1051-INV',
      actId: 'DOC-DEMO-1051-ACT',
      actNumber: 'АНП-1051',
      createdAt: '01.05.2026 17:20',
    } satisfies TaxInvoiceSnapshot),
    createdAt: '01.05.2026 17:20',
    registeredAt: '01.05.2026 17:35',
    responsible: 'Вікторія Данилюк',
  },
];

const basExchangeItems: BasExchangeItem[] = [
  { id: 'BAS-1', platformEntityType: 'counterparty', platformEntityId: 'CL-ALFA', basEntityId: 'BAS-CTR-00031', exchangeStatus: 'imported', payloadJson: '{"type":"counterparty","edrpou":"22590485","name":"ТОВ Альфа-Сервіс"}', sentAt: '19.04.2026 16:10', confirmedAt: '19.04.2026 16:12', retryCount: 0, amountNet: 0, amountVat: 0, amountGross: 0, client: 'ТОВ "Альфа-Сервіс"', createdBy: 'Ірина Бондар', title: 'ТОВ "Альфа-Сервіс"', responsible: 'Ірина Бондар' },
  { id: 'BAS-2', platformEntityType: 'act', platformEntityId: 'ЗН-1047', exchangeStatus: 'queued', payloadJson: '{"type":"act","orderId":"ЗН-1047","client":"ТОВ Альфа-Сервіс","items":["роботи","запчастини"]}', retryCount: 0, amountNet: 2075, amountVat: 415, amountGross: 2490, orderId: 'ЗН-1047', client: 'ТОВ "Альфа-Сервіс"', createdBy: 'Вікторія Данилюк', title: 'Акт виконаних робіт ЗН-1047', responsible: 'Ірина Бондар' },
  { id: 'BAS-3', platformEntityType: 'payment', platformEntityId: 'PAY-1', basEntityId: 'BAS-PAY-00088', exchangeStatus: 'sent', payloadJson: '{"type":"payment","orderId":"ЗН-1047","method":"card","amount":500}', sentAt: '19.04.2026 16:20', retryCount: 0, amountNet: 416.67, amountVat: 83.33, amountGross: 500, orderId: 'ЗН-1047', client: 'ТОВ "Альфа-Сервіс"', createdBy: 'Вікторія Данилюк', title: 'Оплата карткою ЗН-1047', responsible: 'Ірина Бондар' },
  { id: 'BAS-4', platformEntityType: 'tax_invoice', platformEntityId: 'ПН-1047', exchangeStatus: 'new', payloadJson: '{"type":"tax_invoice","source":"act","orderId":"ЗН-1047","vatRate":20}', retryCount: 0, amountNet: 2075, amountVat: 415, amountGross: 2490, orderId: 'ЗН-1047', client: 'ТОВ "Альфа-Сервіс"', createdBy: 'Ірина Бондар', title: 'Чернетка ПН по ЗН-1047', responsible: 'Ірина Бондар' },
  { id: 'BAS-5', platformEntityType: 'purchase', platformEntityId: 'PO-204', exchangeStatus: 'error', payloadJson: '{"type":"purchase","supplier":"Parts UA","items":["CAN-RL-02"]}', retryCount: 2, lastError: 'Не зіставлена номенклатура CAN-RL-02 у BAS.', amountNet: 1208.33, amountVat: 241.67, amountGross: 1450, client: 'Parts UA', createdBy: 'Олена Мороз', title: 'Закупка роликів Canon', responsible: 'Ірина Бондар' },
];

const basMappings: BasMapping[] = [
  { directory: 'Контрагенти', platformValue: 'ТОВ "Альфа-Сервіс"', basValue: 'ТОВ Альфа-Сервіс / BAS-CTR-00031', status: 'Зіставлено' },
  { directory: 'Номенклатура', platformValue: 'Ролик подачі Canon', basValue: 'Потрібно вибрати номенклатуру BAS', status: 'Потрібно зіставити' },
  { directory: 'Послуги', platformValue: 'Ремонт плати живлення', basValue: 'Сервісні роботи / BAS-SRV-001', status: 'Зіставлено' },
  { directory: 'Склади', platformValue: 'Основний склад', basValue: 'Склад сервісу / BAS-WH-01', status: 'Зіставлено' },
  { directory: 'Ставки ПДВ', platformValue: '20%', basValue: 'ПДВ 20%', status: 'Зіставлено' },
];

const basImplementationPhases: BasImplementationPhase[] = [
  {
    phase: 'Етап 1',
    title: 'Запуск операційної платформи',
    result: 'Сервісний центр працює у платформі, менеджери та інженери не залежать від 1С.',
    priority: 'Перша черга',
    done: true,
    items: ['Клієнти', 'Заказы ремонту', 'Призначення інженерів', 'Статуси ремонту', 'Склад і запчастини', 'Оплати', 'Видача заказів', 'Документи', 'Уведомлення', 'Зарплата інженерів', 'Полки / QR'],
  },
  {
    phase: 'Етап 2',
    title: 'Блок бухгалтера всередині платформи',
    result: 'Бухгалтер бачить оплату, акти, ПДВ, ПН, борги і статуси без ручного хаосу.',
    priority: 'Перша черга',
    done: true,
    items: ['Контроль оплат', 'Контроль актів', 'Контроль ПДВ', 'Податкові накладні', 'Контроль боргів', 'Статуси актів і оплат'],
  },
  {
    phase: 'Етап 3',
    title: 'Вигрузка в BAS',
    result: 'BAS отримує дані з платформи через Excel / CSV / JSON без повторного ручного введення.',
    priority: 'Перша черга',
    done: true,
    items: ['Контрагенти', 'Акти виконаних робіт', 'Оплати', 'Продажі / реалізації', 'Закупки', 'Чернетки / дані для ПН'],
  },
  {
    phase: 'Етап 4',
    title: 'Журнал обміну',
    result: 'Видно, що пішло в BAS, що прийнято, що впало з помилкою і що треба повторити.',
    priority: 'Друга черга',
    done: true,
    items: ['Новое', 'Вигружено', 'Помилка', 'Журнал помилок', 'Повторна відправка', 'ID документа BAS', 'Історія обміну'],
  },
  {
    phase: 'Етап 5',
    title: 'Автоматична інтеграція з BAS',
    result: 'Платформа і BAS працюють як єдина система після перевірки стабільності вигрузок.',
    priority: 'Друга черга',
    done: false,
    items: ['Автоотправка актів', 'Автоотправка оплат', 'Автоотправка реалізацій', 'Автоотправка закупок', 'Дані по ПН', 'Зворотні статуси з BAS'],
  },
  {
    phase: 'Етап 6',
    title: 'Зворотний обмін із BAS',
    result: 'Платформа бачить не тільки відправку, а й результат проведення / реєстрації / банківського підтвердження.',
    priority: 'Друга черга',
    done: false,
    items: ['Статус реєстрації ПН', 'Підтвердження завантаження', 'Помилки реєстрації', 'Проведено / не проведено', 'Підтвердження оплати по банку'],
  },
  {
    phase: 'Не старт',
    title: 'Не робити одразу',
    result: 'Не ускладнюємо запуск і не затягуємо операційну платформу.',
    priority: 'Не робити на старті',
    done: false,
    items: ['Повна бухгалтерія всередині платформи', 'ОСВ', 'План рахунків', 'Складні проводки', 'Повний двосторонній обмін з першого дня'],
  },
];

const initialOrderUnits: OrderUnit[] = [
  { id: 'UNIT-1048-1', orderId: 'ЗН-1048', groupId: 'GR-20260419-1', client: 'Олена Коваленко', type: 'Ноутбук', code: 'QR-UNIT-1048-1', status: 'Діагностика', engineer: 'Андрій Савчук', lastActionAt: '19.04.2026 09:35' },
  { id: 'UNIT-1047-1', orderId: 'ЗН-1047', groupId: 'GR-20260418-1', client: 'ТОВ "Альфа-Сервіс"', type: 'МФУ', code: 'QR-UNIT-1047-1', status: 'В ремонті', engineer: 'Марина Лисенко', lastActionAt: '18.04.2026 14:20' },
  { id: 'UNIT-1047-C1', orderId: 'ЗН-1047', groupId: 'GR-20260418-1', client: 'ТОВ "Альфа-Сервіс"', type: 'Картридж', code: 'BAR-1047-C1', status: 'Готово', engineer: 'Марина Лисенко', locationCode: 'A-1-1', lastActionAt: '19.04.2026 15:20' },
  { id: 'UNIT-1047-C2', orderId: 'ЗН-1047', groupId: 'GR-20260418-1', client: 'ТОВ "Альфа-Сервіс"', type: 'Картридж', code: 'BAR-1047-C2', status: 'На полці', engineer: 'Марина Лисенко', locationCode: 'A-1-2', lastActionAt: '19.04.2026 15:30' },
];

const storageCells: StorageCell[] = [
  { id: 'CELL-A-1-1', shelf: 'A', level: 1, position: 1, code: 'A-1-1', occupiedBy: 'UNIT-1047-C1' },
  { id: 'CELL-A-1-2', shelf: 'A', level: 1, position: 2, code: 'A-1-2', occupiedBy: 'UNIT-1047-C2' },
  { id: 'CELL-A-1-3', shelf: 'A', level: 1, position: 3, code: 'A-1-3' },
  { id: 'CELL-A-2-1', shelf: 'A', level: 2, position: 1, code: 'A-2-1' },
  { id: 'CELL-B-1-1', shelf: 'B', level: 1, position: 1, code: 'B-1-1' },
  { id: 'CELL-B-1-2', shelf: 'B', level: 1, position: 2, code: 'B-1-2' },
];

const initialWarehouseLocations: WarehouseLocation[] = [
  { id: 'LOC-A-3-1-1', zone: 'REPAIR', rack: 'A', shelf: '3-1-1', code: 'A-3-1-1' },
  { id: 'LOC-A-3-1-2', zone: 'REPAIR', rack: 'A', shelf: '3-1-2', code: 'A-3-1-2' },
  { id: 'LOC-A-3-1-3', zone: 'REPAIR', rack: 'A', shelf: '3-1-3', code: 'A-3-1-3' },
  { id: 'LOC-B-2-1-1', zone: 'REPAIR', rack: 'B', shelf: '2-1-1', code: 'B-2-1-1' },
  { id: 'LOC-B-2-1-2', zone: 'REPAIR', rack: 'B', shelf: '2-1-2', code: 'B-2-1-2' },
  { id: 'LOC-B-2-1-3', zone: 'REPAIR', rack: 'B', shelf: '2-1-3', code: 'B-2-1-3' },
  { id: 'LOC-C-1-1-1', zone: 'STORAGE', rack: 'C', shelf: '1-1-1', code: 'C-1-1-1' },
  { id: 'LOC-C-1-1-2', zone: 'STORAGE', rack: 'C', shelf: '1-1-2', code: 'C-1-1-2' },
];

const acceptanceChecklist = [
  { group: 'Загальне', item: 'Система працює через браузер і не привʼязана до одного ПК', done: true },
  { group: 'Ролі', item: 'Інженер бачить тільки свої замовлення', done: true },
  { group: 'Маркування', item: 'QR/штрихкод відкриває заказ або одиницю', done: true },
  { group: 'Полки', item: 'При статусі На полці місце зберігання обовʼязкове', done: true },
  { group: 'Зарплата', item: 'Ремонт рахується від робіт, картриджі за фікс за одиницю', done: true },
  { group: 'НДС', item: 'Акт видано → готово до реєстрації ПН, переносу періоду немає', done: true },
  { group: 'Безпека', item: 'Закриті документи не редагуються без права адміністратора', done: true },
  { group: 'Backup', item: 'Щоденний backup потрібен на VPS/Cloud етапі', done: false },
];

const initialClients: ClientRecord[] = [
  { name: 'Олена Коваленко', phone: '+380 67 441 28 10', taxId: '', orders: 3 },
  { name: 'ТОВ "Альфа-Сервіс"', phone: '+380 44 301 88 20', taxId: '41234567', orders: 12 },
  { name: 'Ігор Мельник', phone: '+380 50 617 73 04', taxId: '', orders: 2 },
  { name: 'Наталія Романюк', phone: '+380 93 118 45 66', taxId: '', orders: 5 },
  { name: 'ТОВ "Демо ПДВ"', phone: '+380 44 555 10 50', taxId: '44556677', orders: 1 },
  { name: 'ТОВ "TEST ПН"', phone: '+380 44 555 10 51', taxId: '44556678', orders: 1 },
];

const payrollRules: PayrollRule[] = [
  { employee: 'Андрій Савчук', role: 'Інженер', type: 'Процент від роботи', percent: 40 },
  { employee: 'Марина Лисенко', role: 'Інженер', type: 'Процент від роботи', percent: 25 },
  { employee: 'Олег Шевченко', role: 'Інженер', type: 'Процент від роботи', percent: 35 },
  { employee: 'Олена Мороз', role: 'Менеджер', type: 'Фіксована ставка', monthlyRate: 15000, bonus: 1000 },
  { employee: 'Вікторія Данилюк', role: 'Адміністратор', type: 'Фіксована ставка', monthlyRate: 18000, bonus: 1500 },
];

const EMPLOYEES_STORAGE_KEY = 'crm_employees';
const CLIENTS_STORAGE_KEY = 'crm_clients';
const PRODUCTS_STORAGE_KEY = 'crm_products';
const ORDERS_STORAGE_KEY = 'crm_orders';
const PAYMENTS_STORAGE_KEY = 'crm_payments';
const TAX_INVOICES_STORAGE_KEY = 'crm_tax_invoices';
const CONTRACTS_STORAGE_KEY = 'crm_contracts';
const CONTRACT_ACTS_STORAGE_KEY = 'crm_contract_acts';
const BANK_IMPORTS_STORAGE_KEY = 'crm_bank_imports';
const REQUIREMENTS_STORAGE_KEY = 'crm_requirements';
const PURCHASES_STORAGE_KEY = 'crm_purchases';
const RECEIPTS_STORAGE_KEY = 'crm_receipts';
const MOVEMENTS_STORAGE_KEY = 'crm_movements';
const ACTION_LOGS_STORAGE_KEY = 'crm_action_logs';
const BACKUPS_STORAGE_KEY = 'crm_backups';
const CASH_SHIFT_STORAGE_KEY = 'crm_cash_shift';
const MAX_ACTION_LOGS = 100;

function loadEmployeesFromStorage(): User[] {
  try {
    const raw = localStorage.getItem(EMPLOYEES_STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? normalizeUsers(parsed as User[]) : [];
  } catch {
    return [];
  }
}

function saveEmployeesToStorage(employees: User[]) {
  localStorage.setItem(EMPLOYEES_STORAGE_KEY, JSON.stringify(employees));
}

function mergeSeededById<T extends { id: string }>(stored: T[], seeded: T[]) {
  return [...seeded.filter((item) => !stored.some((current) => current.id === item.id)), ...stored];
}

function mergeSeededClients(stored: ClientRecord[], seeded: ClientRecord[]) {
  return [
    ...seeded.filter((seed) => !stored.some((client) => (
      extractDigits(client.phone) === extractDigits(seed.phone)
      || normalizeLooseText(client.name) === normalizeLooseText(seed.name)
    ))),
    ...stored,
  ];
}

function loadClientsFromStorage(): ClientRecord[] {
  try {
    const raw = localStorage.getItem(CLIENTS_STORAGE_KEY);
    if (!raw) return initialClients;
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed)
      ? mergeSeededClients(parsed.map((client) => ({
          name: String(client.name ?? '').trim(),
          phone: normalizeImportedPhone(client.phone),
          email: client.email ? String(client.email).trim() : undefined,
          taxId: normalizeTaxId(client.taxId),
          orders: Number(client.orders) || 0,
        } satisfies ClientRecord)), initialClients)
      : initialClients;
  } catch {
    return initialClients;
  }
}

function saveClientsToStorage(clients: ClientRecord[]) {
  localStorage.setItem(CLIENTS_STORAGE_KEY, JSON.stringify(clients.map((client) => ({
    ...client,
    taxId: normalizeTaxId(client.taxId),
  }))));
}

function loadProductsFromStorage(): Product[] {
  try {
    const raw = localStorage.getItem(PRODUCTS_STORAGE_KEY);
    if (!raw) return initialProducts.map((product) => normalizeStoredProduct(product));
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed.map((product) => normalizeStoredProduct(product as Product)) : initialProducts.map((product) => normalizeStoredProduct(product));
  } catch {
    return initialProducts.map((product) => normalizeStoredProduct(product));
  }
}

function saveProductsToStorage(products: Product[]) {
  localStorage.setItem(PRODUCTS_STORAGE_KEY, JSON.stringify(products.map((product) => normalizeStoredProduct(product))));
}

function loadOrdersFromStorage(): ServiceOrder[] {
  try {
    const raw = localStorage.getItem(ORDERS_STORAGE_KEY);
    if (!raw) return initialOrders;
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? mergeSeededById(parsed as ServiceOrder[], initialOrders) : initialOrders;
  } catch {
    return initialOrders;
  }
}

function saveOrdersToStorage(orders: ServiceOrder[]) {
  localStorage.setItem(ORDERS_STORAGE_KEY, JSON.stringify(orders));
}

function parseTaxInvoiceSnapshot(snapshot?: string) {
  try {
    if (!snapshot) return null;
    const parsed = JSON.parse(snapshot);
    return parsed && typeof parsed === 'object' ? parsed as TaxInvoiceSnapshot : null;
  } catch {
    return null;
  }
}

function normalizeTaxInvoiceStatus(value?: string): TaxInvoiceStatus {
  if (value === 'Зареєстровано') return 'Зареєстровано';
  if (value === 'Помилка' || value === 'Зупинено' || value === 'Прострочено') return 'Помилка';
  if (value === 'Очікує ПН') return 'Очікує ПН';
  return 'Створено';
}

function normalizeStoredTaxInvoice(invoice: Partial<TaxInvoice> & Record<string, unknown>): TaxInvoice {
  const legacyAmount = Number(invoice.amount ?? 0) || 0;
  const legacyVatAmount = Number(invoice.vatAmount ?? 0) || 0;
  const legacyFormedAt = typeof invoice['formedAt'] === 'string' ? invoice['formedAt'] : undefined;
  const eventDate = String(invoice.date ?? invoice.eventDate ?? (extractDayKey(String(invoice.createdAt ?? legacyFormedAt ?? today)) || today));
  const snapshot = parseTaxInvoiceSnapshot(typeof invoice.snapshot === 'string' ? invoice.snapshot : '');
  const normalizedSnapshot: TaxInvoiceSnapshot = snapshot ?? {
    companyName: companyRequisites.name,
    companyEdrpou: companyRequisites.edrpou,
    companyIpn: companyRequisites.ipn,
    companyIban: companyRequisites.iban,
    companyBank: companyRequisites.bank,
    companyMfo: companyRequisites.mfo,
    companyAddress: companyRequisites.address,
    companyPhone: companyRequisites.phone,
    companyVatStatus: 'Платник ПДВ',
    clientName: String(invoice.client ?? 'Невідомий клієнт'),
    clientPhone: '',
    clientTaxId: '',
    orderId: String(invoice.orderId ?? ''),
    orderNumber: String(invoice.orderId ?? ''),
    device: '',
    managerName: String(invoice.responsible ?? SYSTEM_USER.name),
    documentNumber: String(invoice.number ?? invoice.id ?? ''),
    documentDate: eventDate,
    eventDate,
    eventType: (invoice.eventType === 'payment' ? 'payment' : 'act'),
    sumWithoutVAT: Math.round((legacyAmount - legacyVatAmount) * 100) / 100,
    vatAmount: legacyVatAmount,
    sumWithVAT: legacyAmount,
    totalWithVATWords: numberToWordsUA(legacyAmount),
    vatAmountWords: numberToWordsUA(legacyVatAmount),
    lines: [],
    paymentId: typeof invoice.paymentId === 'string' ? invoice.paymentId : undefined,
    paymentType: undefined,
    paymentStatus: undefined,
    actId: typeof invoice.actId === 'string' ? invoice.actId : undefined,
    createdAt: String(invoice.createdAt ?? legacyFormedAt ?? today),
  };
  return {
    id: String(invoice.id ?? uid('TAX')),
    number: String(invoice.number ?? invoice.id ?? uid('ПН')),
    date: eventDate,
    orderId: String(invoice.orderId ?? normalizedSnapshot.orderId ?? ''),
    invoiceId: typeof invoice.invoiceId === 'string' ? invoice.invoiceId : normalizedSnapshot.invoiceId,
    actId: typeof invoice.actId === 'string' ? invoice.actId : normalizedSnapshot.actId,
    paymentId: typeof invoice.paymentId === 'string' ? invoice.paymentId : normalizedSnapshot.paymentId,
    waybillId: typeof invoice.waybillId === 'string' ? invoice.waybillId : normalizedSnapshot.waybillId,
    eventType: invoice.eventType === 'payment' || invoice.eventType === 'act' ? invoice.eventType : normalizedSnapshot.eventType,
    client: String(invoice.client ?? normalizedSnapshot.clientName ?? ''),
    eventDate,
    amount: Number(invoice.amount ?? normalizedSnapshot.sumWithVAT ?? 0) || 0,
    vatAmount: Number(invoice.vatAmount ?? normalizedSnapshot.vatAmount ?? 0) || 0,
    status: normalizeTaxInvoiceStatus(String(invoice.status ?? 'Створено')),
    snapshot: JSON.stringify(normalizedSnapshot),
    createdAt: String(invoice.createdAt ?? legacyFormedAt ?? normalizedSnapshot.createdAt ?? today),
    registeredAt: typeof invoice.registeredAt === 'string' ? invoice.registeredAt : undefined,
    responsible: String(invoice.responsible ?? normalizedSnapshot.managerName ?? SYSTEM_USER.name),
  };
}

function loadTaxInvoicesFromStorage(): TaxInvoice[] {
  try {
    const raw = localStorage.getItem(TAX_INVOICES_STORAGE_KEY);
    if (!raw) return initialTaxInvoices.map(normalizeStoredTaxInvoice);
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed)
      ? mergeSeededById(
          parsed.map((item) => normalizeStoredTaxInvoice(item as Partial<TaxInvoice> & Record<string, unknown>)),
          initialTaxInvoices.map(normalizeStoredTaxInvoice),
        )
      : initialTaxInvoices.map(normalizeStoredTaxInvoice);
  } catch {
    return initialTaxInvoices.map(normalizeStoredTaxInvoice);
  }
}

function saveTaxInvoicesToStorage(items: TaxInvoice[]) {
  localStorage.setItem(TAX_INVOICES_STORAGE_KEY, JSON.stringify(items));
}

function loadSimplePaymentsFromStorage(): SimpleOrderPaymentRecord[] {
  try {
    const raw = localStorage.getItem(PAYMENTS_STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed)
      ? parsed.map((payment) => ({
          id: payment.id,
          client: payment.client ?? '',
          entityType: payment.entityType ?? 'order',
          entityId: payment.entityId ?? payment.orderId ?? payment.contractId ?? payment.actId ?? '',
          clientTaxId: normalizeTaxId(payment.clientTaxId ?? payment.taxId),
          orderId: payment.orderId,
          contractId: payment.contractId,
          actId: payment.actId,
          amount: Number(payment.amount) || 0,
          method: payment.method ?? payment.type ?? 'наличные',
          paymentKind: payment.paymentKind ?? 'оплата',
          status: payment.status ?? 'подтвержден',
          date: payment.date ?? today,
          currency: payment.currency,
          payer: payment.payer,
          taxId: normalizeTaxId(payment.taxId),
          purpose: payment.purpose,
          documentRef: payment.documentRef,
          direction: payment.direction ?? 'incoming',
          bankAccount: payment.bankAccount,
          bankName: payment.bankName,
          fingerprint: payment.fingerprint,
        } satisfies SimpleOrderPaymentRecord))
      : [];
  } catch {
    return [];
  }
}

function saveSimplePaymentsToStorage(payments: SimpleOrderPaymentRecord[]) {
  localStorage.setItem(PAYMENTS_STORAGE_KEY, JSON.stringify(payments));
}

function loadContractsFromStorage() {
  try {
    const raw = localStorage.getItem(CONTRACTS_STORAGE_KEY);
    if (!raw) return initialContracts;
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed as ContractRecord[] : initialContracts;
  } catch {
    return initialContracts;
  }
}

function saveContractsToStorage(contracts: ContractRecord[]) {
  localStorage.setItem(CONTRACTS_STORAGE_KEY, JSON.stringify(contracts));
}

function loadContractActsFromStorage() {
  try {
    const raw = localStorage.getItem(CONTRACT_ACTS_STORAGE_KEY);
    if (!raw) return [] as ContractActRecord[];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed)
      ? parsed.map((act) => {
          const amount = Number(act.amount) || 0;
          const paidAmount = Math.max(0, Number(act.paidAmount) || (act.status === 'Оплачено' ? amount : 0));
          const remainingAmount = Math.max(amount - paidAmount, 0);
          return {
            id: act.id,
            contractId: act.contractId,
            orderIds: Array.isArray(act.orderIds) ? act.orderIds : [],
            amount,
            paidAmount,
            remainingAmount,
            date: act.date ?? today,
            status: remainingAmount <= 0 && amount > 0 ? 'Оплачен' : paidAmount > 0 ? 'Частично оплачен' : 'Не оплачено',
          } satisfies ContractActRecord;
        })
      : [];
  } catch {
    return [] as ContractActRecord[];
  }
}

function saveContractActsToStorage(acts: ContractActRecord[]) {
  localStorage.setItem(CONTRACT_ACTS_STORAGE_KEY, JSON.stringify(acts));
}

function loadBankImportItemsFromStorage() {
  try {
    const raw = localStorage.getItem(BANK_IMPORTS_STORAGE_KEY);
    if (!raw) return [] as BankImportItem[];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed as BankImportItem[] : [];
  } catch {
    return [] as BankImportItem[];
  }
}

function saveBankImportItemsToStorage(items: BankImportItem[]) {
  localStorage.setItem(BANK_IMPORTS_STORAGE_KEY, JSON.stringify(items));
}

function loadRequirementsFromStorage() {
  try {
    const raw = localStorage.getItem(REQUIREMENTS_STORAGE_KEY);
    if (!raw) return [] as PartRequirement[];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed as PartRequirement[] : [];
  } catch {
    return [] as PartRequirement[];
  }
}

function saveRequirementsToStorage(items: PartRequirement[]) {
  localStorage.setItem(REQUIREMENTS_STORAGE_KEY, JSON.stringify(items));
}

function loadPurchasesFromStorage() {
  try {
    const raw = localStorage.getItem(PURCHASES_STORAGE_KEY);
    if (!raw) return [] as PurchaseOrder[];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed as PurchaseOrder[] : [];
  } catch {
    return [] as PurchaseOrder[];
  }
}

function savePurchasesToStorage(items: PurchaseOrder[]) {
  localStorage.setItem(PURCHASES_STORAGE_KEY, JSON.stringify(items));
}

function loadReceiptsFromStorage() {
  try {
    const raw = localStorage.getItem(RECEIPTS_STORAGE_KEY);
    if (!raw) return [] as GoodsReceipt[];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed as GoodsReceipt[] : [];
  } catch {
    return [] as GoodsReceipt[];
  }
}

function saveReceiptsToStorage(items: GoodsReceipt[]) {
  localStorage.setItem(RECEIPTS_STORAGE_KEY, JSON.stringify(items));
}

function loadMovementsFromStorage() {
  try {
    const raw = localStorage.getItem(MOVEMENTS_STORAGE_KEY);
    if (!raw) return [] as StockMovement[];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed as StockMovement[] : [];
  } catch {
    return [] as StockMovement[];
  }
}

function saveMovementsToStorage(items: StockMovement[]) {
  localStorage.setItem(MOVEMENTS_STORAGE_KEY, JSON.stringify(items));
}

function limitActionLogs(items: ActionLog[]) {
  return items.slice(0, MAX_ACTION_LOGS);
}

function loadActionLogsFromStorage() {
  return [] as ActionLog[];
}

function saveActionLogsToStorage(_items: ActionLog[]) {
  // Temporarily disabled to prevent QuotaExceededError from crm_action_logs.
}

function loadCashShiftFromStorage() {
  try {
    const raw = localStorage.getItem(CASH_SHIFT_STORAGE_KEY);
    if (!raw) return initialCashShift;
    const parsed = JSON.parse(raw) as Partial<CashShift>;
    return {
      ...initialCashShift,
      ...parsed,
    } as CashShift;
  } catch {
    return initialCashShift;
  }
}

function saveCashShiftToStorage(item: CashShift) {
  localStorage.setItem(CASH_SHIFT_STORAGE_KEY, JSON.stringify(item));
}

function loadBackupsFromStorage() {
  try {
    const raw = localStorage.getItem(BACKUPS_STORAGE_KEY);
    if (!raw) return [] as BackupSnapshot[];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed as BackupSnapshot[] : [];
  } catch {
    return [] as BackupSnapshot[];
  }
}

function saveBackupsToStorage(items: BackupSnapshot[]) {
  localStorage.setItem(BACKUPS_STORAGE_KEY, JSON.stringify(items));
}

const SYSTEM_USER: User = {
  id: 'system',
  name: 'System',
  role: 'Адміністратор',
  twoFactor: false,
  login: 'system',
  authMode: 'phone_code',
  permissions: {
    canAccessWarehouse: true,
    canAccessFinance: true,
    canAccessEmployees: true,
    canAccessSettings: true,
    canAccessReports: true,
  },
  session: 'Активна',
};

const companyBankAccounts: BankAccountRecord[] = [
  { id: 'BANK-RAIF-1', bankName: 'Райффайзен', iban: 'UA123808050000000002600000001', currency: 'UAH', legalType: 'ТОВ' },
  { id: 'BANK-PRIVAT-1', bankName: 'ПриватБанк', iban: 'UA213052990000026005000000002', currency: 'UAH', legalType: 'ФОП' },
];

const initialInternalMessages: InternalMessage[] = [
  {
    id: 'MSG-IN-1',
    toUser: 'Олена Мороз',
    toRole: 'Менеджер',
    orderId: 'ЗН-1047',
    text: 'Перевірити оплату і підписаний акт перед видачею.',
    dueAt: '20.04.2026 17:00',
    importance: 'Важлива',
    status: 'Нове',
    createdBy: 'Сергій Павловський',
    createdAt: today,
  },
  {
    id: 'MSG-IN-2',
    toUser: 'Андрій Савчук',
    toRole: 'Інженер',
    orderId: 'ЗН-1048',
    text: 'Додати причину діагностики або поставити статус "Не підлягає ремонту", якщо плата не відновлюється.',
    dueAt: '20.04.2026 15:00',
    importance: 'Звичайна',
    status: 'Нове',
    createdBy: 'Сергій Павловський',
    createdAt: today,
  },
  {
    id: 'MSG-IN-3',
    toUser: 'Ірина Бондар',
    toRole: 'Бухгалтер',
    orderId: 'ЗН-1047',
    text: 'Підготувати ПН після видачі акта, оплату контролюємо окремо.',
    dueAt: '20.04.2026 18:00',
    importance: 'Критична',
    status: 'Прочитано',
    createdBy: 'Сергій Павловський',
    createdAt: today,
  },
];

const rolePermissions: Record<Role, string[]> = {
  Руководитель: ['Повний контроль', 'Фінанси', 'Звіти', 'Ролі', 'Користувачі', 'Експорт'],
  Адміністратор: ['Ремонти', 'Продажі', 'Оплати', 'Повернення', 'Клієнти', 'Документи'],
  Менеджер: ['Створення ремонтів', 'Створення продажів', 'Резерв', 'Оплата', 'Видача'],
  Інженер: ['Перегляд ремонтів', 'Встановлення запчастин', 'Повернення деталі в сервісі'],
  Комірник: ['Залишки', 'Прихід', 'Резерв', 'Списання', 'Інвентаризація', 'Коригування'],
  Бухгалтер: ['Каса', 'Платежі', 'Повернення коштів', 'Звіти', 'НДС', 'Закупівлі', 'Собівартість', 'Експорт'],
  Закупник: ['Закупівлі', 'Постачальники', 'Статуси закупок', 'Закриття закупки'],
};

const baseEmployeeRoles: Role[] = ['Адміністратор', 'Руководитель', 'Менеджер', 'Інженер', 'Бухгалтер'];

const rolePageAccess: Record<Role, Page[]> = {
  Руководитель: navItems.map((item) => item.id),
  Адміністратор: navItems.map((item) => item.id),
  Менеджер: ['dashboard', 'clients', 'orders'],
  Інженер: ['my-orders'],
  Комірник: ['parts', 'purchases', 'storage', 'movements'],
  Бухгалтер: ['finance', 'cash', 'documents', 'bank-import', 'reports', 'tax-invoices'],
  Закупник: ['parts', 'purchases', 'storage', 'movements'],
};

const roleFinePermissions: Record<Role, Permission[]> = {
  Руководитель: [
    'clients.view',
    'repairs.view',
    'sales.view',
    'stock.view',
    'purchases.view',
    'finance.cash', 'finance.reports',
    'reports.view', 'reports.export', 'reports.profit', 'reports.staff',
    'cost.view', 'profit.view',
  ],
  Адміністратор: [
    'clients.view', 'clients.create', 'clients.edit',
    'repairs.view', 'repairs.create', 'repairs.edit', 'repairs.status', 'repairs.close', 'repairs.cancel',
    'sales.view', 'sales.create', 'sales.edit', 'sales.post', 'sales.cancel', 'sales.return',
    'stock.view', 'stock.receive', 'stock.reserve', 'stock.unreserve', 'stock.writeoff', 'stock.inventory', 'stock.adjust',
    'purchases.view', 'purchases.create', 'purchases.edit', 'purchases.send', 'purchases.status', 'purchases.close',
    'finance.cash', 'finance.payments', 'finance.reports', 'finance.refund', 'finance.adjust',
    'reports.view', 'reports.export', 'reports.profit', 'reports.staff',
    'settings.roles', 'settings.users', 'settings.templates', 'settings.directories', 'settings.system',
    'cost.view', 'profit.view', 'discount.approve.high', 'documents.delete.posted', 'documents.edit.closed', 'payments.backdate.edit',
  ],
  Менеджер: [
    'clients.view', 'clients.create', 'clients.edit',
    'repairs.view', 'repairs.create', 'repairs.edit', 'repairs.status',
    'sales.view', 'sales.create', 'sales.edit', 'sales.post', 'sales.cancel', 'sales.return',
    'finance.payments',
  ],
  Інженер: [
    'repairs.view', 'repairs.edit', 'repairs.status',
  ],
  Комірник: [
    'stock.view', 'stock.receive', 'stock.reserve', 'stock.unreserve', 'stock.writeoff', 'stock.inventory', 'stock.adjust',
    'purchases.view', 'purchases.create', 'purchases.edit', 'purchases.send', 'purchases.status',
    'sales.view',
  ],
  Бухгалтер: [
    'clients.view',
    'repairs.view',
    'sales.view',
    'stock.view',
    'purchases.view',
    'finance.cash', 'finance.payments', 'finance.reports', 'finance.refund', 'finance.adjust',
    'reports.view', 'reports.export', 'reports.profit',
    'cost.view', 'profit.view',
  ],
  Закупник: [
    'stock.view',
    'purchases.view', 'purchases.create', 'purchases.edit', 'purchases.send', 'purchases.status', 'purchases.close',
    'reports.view',
    'cost.view',
  ],
};

const initialCashShift: CashShift = {
  id: 'SHIFT-CLOSED',
  openedAt: '',
  openedBy: '',
  status: 'Закрита',
  openingCash: 0,
  cashIncome: 0,
  cardIncome: 0,
  bankIncome: 0,
  cashExpense: 0,
};

const statusClass: Record<OrderPartStatus | OrderStatus | RequirementStatus | PurchaseStatus | SaleStatus, string> = {
  'Нова': 'tag tag-amber',
  'В роботі': 'tag tag-blue',
  'Не потрібно': 'tag tag-gray',
  Потрібно: 'tag tag-amber',
  'До закупівлі': 'tag tag-amber',
  Замовлено: 'tag tag-blue',
  'В дорозі': 'tag tag-blue',
  Прибуло: 'tag tag-teal',
  Отримано: 'tag tag-teal',
  Зарезервовано: 'tag tag-green',
  'Видано інженеру': 'tag tag-green',
  Встановлено: 'tag tag-teal',
  Списано: 'tag tag-dark',
  Повернення: 'tag tag-gray',
  Прийнято: 'tag tag-blue',
  'На діагностиці': 'tag tag-amber',
  'Очікує погодження': 'tag tag-amber',
  Погоджено: 'tag tag-green',
  'Очікує запчастину': 'tag tag-gray',
  'В ремонті': 'tag tag-green',
  'На тестуванні': 'tag tag-blue',
  'Готовий до видачі': 'tag tag-teal',
  'Очікує оплати': 'tag tag-amber',
  Видано: 'tag tag-teal',
  Закрито: 'tag tag-dark',
  Скасовано: 'tag tag-gray',
  'Не підлягає ремонту': 'tag tag-dark',
  'Очікує клієнта': 'tag tag-amber',
  'Пауза / відкладено': 'tag tag-gray',
  Створено: 'tag tag-gray',
  'Частково прибуло': 'tag tag-amber',
  Чернетка: 'tag tag-gray',
  'Частково оплачено': 'tag tag-amber',
  Оплачено: 'tag tag-green',
};

function money(value: number) {
  return new Intl.NumberFormat('uk-UA', { style: 'currency', currency: 'UAH', maximumFractionDigits: 0 }).format(value);
}

function uid(prefix: string) {
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`;
}

function available(product: Product) {
  return Math.max(product.stock - product.reserved, 0);
}

function batchTotalAvailable(product: Product) {
  return product.batches.reduce((sum, batch) => sum + batch.qtyAvailable, 0);
}

function oldestBatchAge(product: Product) {
  if (product.batches.length === 0) return 0;
  return Math.max(...product.batches.map((batch) => daysSince(batch.purchaseDate)));
}

function averagePurchasePrice(product: Product) {
  const qtyTotal = product.batches.reduce((sum, batch) => sum + batch.qtyTotal, 0);
  if (qtyTotal <= 0) return product.cost;
  return Math.round(product.batches.reduce((sum, batch) => sum + (batch.purchasePrice * batch.qtyTotal), 0) / qtyTotal);
}

function normalizeImportHeader(value: unknown) {
  return String(value ?? '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function normalizeImportedPhone(value: unknown) {
  return String(value ?? '').trim();
}

function parseImportedAmount(value: unknown) {
  const normalized = String(value ?? '').replace(/\s+/g, '').replace(',', '.').replace(/[^\d.-]/g, '');
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

function normalizeTaxId(value: unknown) {
  return String(value ?? '').replace(/\D/g, '');
}

function normalizePaymentText(value: unknown) {
  return String(value ?? '')
    .toLowerCase()
    .replace(/[^\p{L}\p{N}\s]/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function extractReferenceTokens(value: unknown) {
  const normalized = normalizePaymentText(value);
  return Array.from(new Set(normalized.match(/[a-zа-яіїє0-9-]{3,}/gu) ?? []));
}

function bankImportFingerprint(row: Pick<BankStatementRow, 'date' | 'amount' | 'payer' | 'purpose' | 'documentRef' | 'bankAccount'>) {
  return [
    normalizePaymentText(row.date),
    row.amount.toFixed(2),
    normalizePaymentText(row.payer),
    normalizePaymentText(row.purpose),
    normalizePaymentText(row.documentRef),
    normalizePaymentText(row.bankAccount),
  ].join('|');
}

const bankImportFieldAliases: Record<BankImportField, string[]> = {
  date: ['дата', 'date', 'дата платежу', 'дата платежа', 'дата операції', 'дата операции'],
  amount: ['сума', 'сумма', 'amount', 'credit', 'надходження', 'зарахування'],
  currency: ['валюта', 'currency', 'curr'],
  payer: ['платник', 'плательщик', 'payer', 'контрагент', 'назва контрагента', 'наименование контрагента'],
  taxId: ['єдрпоу', 'едрпоу', 'інн', 'инн', 'код', 'єдрпо', 'edrpou', 'tax id'],
  purpose: ['призначення платежу', 'назначение платежа', 'purpose', 'description', 'деталі операції', 'детали операции'],
  documentRef: ['номер документа', 'номер рахунку', 'номер счета', 'документ', 'рахунок', 'счет', 'акт', 'invoice', 'номер операції', 'номер операции', 'ref'],
  direction: ['тип операції', 'тип операции', 'напрямок', 'направление', 'debit/credit', 'debit credit', 'вхідний/вихідний'],
  bankAccount: ['iban', 'рахунок', 'счет', 'iban отримувача', 'iban получателя', 'номер рахунку', 'номер счета', 'account'],
};

function detectBankImportMapping(rows: Record<string, unknown>[]) {
  const headers = Object.keys(rows[0] ?? {});
  const normalizedHeaders = headers.map((header) => normalizeImportHeader(header));
  const mapping: Partial<Record<BankImportField, string>> = {};
  (Object.keys(bankImportFieldAliases) as BankImportField[]).forEach((field) => {
    const alias = bankImportFieldAliases[field].find((candidate) => normalizedHeaders.includes(candidate));
    if (!alias) return;
    const headerIndex = normalizedHeaders.indexOf(alias);
    mapping[field] = headers[headerIndex];
  });
  const detectedBank = normalizedHeaders.some((header) => header.includes('райфф')) ? 'Райффайзен'
    : normalizedHeaders.some((header) => header.includes('приват')) ? 'ПриватБанк'
      : normalizedHeaders.some((header) => header.includes('monobank')) ? 'monobank'
        : 'Невідомий формат';
  const isRecognized = Boolean(mapping.date && mapping.amount && mapping.payer);
  return { mapping, headers, detectedBank, isRecognized };
}

function resolveMappedValue(row: Record<string, unknown>, field: BankImportField, mapping: Partial<Record<BankImportField, string>>) {
  const header = mapping[field];
  if (!header) return '';
  return row[header] ?? '';
}

function detectBankDirection(value: unknown): BankDirection {
  const normalized = String(value ?? '').trim().toLowerCase();
  if (!normalized) return 'unknown';
  if (['вхідний', 'входящий', 'incoming', 'credit', 'надходження', 'зарахування'].some((token) => normalized.includes(token))) return 'incoming';
  if (['вихідний', 'исходящий', 'outgoing', 'debit', 'списання'].some((token) => normalized.includes(token))) return 'outgoing';
  return 'unknown';
}

function readSpreadsheetRows(file: File): Promise<Record<string, unknown>[]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const result = reader.result;
        if (typeof result !== 'string' && !(result instanceof ArrayBuffer)) {
          reject(new Error('Не вдалося прочитати файл.'));
          return;
        }
        const lowerName = file.name.toLowerCase();
        if (lowerName.endsWith('.csv')) {
          const workbook = XLSX.read(result, { type: 'string' });
          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          resolve(XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '' }));
          return;
        }
        if (lowerName.endsWith('.dbf')) {
          const workbook = XLSX.read(result, { type: 'array' });
          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          resolve(XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '' }));
          return;
        }
        const workbook = XLSX.read(result, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        resolve(XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '' }));
      } catch (error) {
        reject(error instanceof Error ? error : new Error('Помилка імпорту файлу.'));
      }
    };
    reader.onerror = () => reject(new Error('Не вдалося прочитати файл.'));
    if (file.name.toLowerCase().endsWith('.csv')) reader.readAsText(file, 'utf-8');
    else reader.readAsArrayBuffer(file);
  });
}

function valueFromRow(row: Record<string, unknown>, aliases: string[]) {
  const entries = Object.entries(row);
  for (const alias of aliases) {
    const match = entries.find(([key]) => normalizeImportHeader(key) === alias);
    if (match) return match[1];
  }
  return '';
}

function importClientsFromRows(rows: Record<string, unknown>[]) {
  return rows
    .map((row) => ({
      name: String(valueFromRow(row, ['імя', "ім'я", 'имя', 'name', 'клієнт', 'клиент']) ?? '').trim(),
      phone: normalizeImportedPhone(valueFromRow(row, ['телефон', 'phone', 'mobile', 'номер'])),
      email: String(valueFromRow(row, ['email', 'e-mail', 'пошта', 'почта']) ?? '').trim() || undefined,
      taxId: normalizeTaxId(valueFromRow(row, ['єдрпоу', 'едрпоу', 'інн', 'инн', 'код', 'tax id', 'edrpou'])),
      orders: 0,
    }))
    .filter((client) => client.name || client.phone || client.email);
}

function normalizeWarehouseText(value: string) {
  return value.trim().replace(/\s+/g, ' ');
}

function normalizeWarehouseKey(value: string) {
  return normalizeWarehouseText(value).toLowerCase();
}

function normalizeBarcodeValue(value: string) {
  return value.trim().replace(/\s+/g, '');
}

function normalizeBarcodeList(...values: Array<string | undefined>) {
  return Array.from(new Set(
    values
      .flatMap((value) => String(value ?? '').split(/[,\n;|]/g))
      .map((item) => normalizeBarcodeValue(item))
      .filter(Boolean),
  ));
}

function buildNormalizedProductName(input: { category?: string; brand?: string; model?: string; name?: string }) {
  const category = normalizeWarehouseText(input.category ?? '');
  const brand = normalizeWarehouseText(input.brand ?? '');
  const model = normalizeWarehouseText(input.model ?? '');
  const fallback = normalizeWarehouseText(input.name ?? '');
  const merged = [category, brand, model].filter(Boolean).join(' ');
  return merged || fallback;
}

function productBarcodeList(product: Product) {
  return normalizeBarcodeList(product.barcode, ...(product.extraBarcodes ?? []));
}

function productMatchesLookup(product: Product, query: string) {
  const normalized = normalizeWarehouseKey(query);
  const barcode = normalizeBarcodeValue(query);
  if (!normalized && !barcode) return false;
  if (normalizeWarehouseKey(product.sku) === normalized) return true;
  if (productBarcodeList(product).some((item) => item === barcode)) return true;
  return [product.name, product.brand, product.model, product.category].some((value) => normalizeWarehouseKey(value).includes(normalized));
}

function normalizeStoredProduct(product: Product): Product {
  const normalizedName = buildNormalizedProductName(product);
  const normalizedSku = normalizeWarehouseText(product.sku ?? '');
  const primaryBarcode = normalizeBarcodeValue(product.barcode || product.sku || '');
  const extraBarcodes = normalizeBarcodeList(...(product.extraBarcodes ?? []), primaryBarcode).filter((item) => item !== primaryBarcode);
  const batches = Array.isArray(product.batches) ? product.batches.map((batch) => ({
    ...batch,
    supplier: batch.supplier ?? batch.source,
    documentNo: batch.documentNo ?? '',
  })) : [];
  return {
    ...product,
    sku: normalizedSku || product.id,
    name: normalizedName || product.name || product.sku,
    barcode: primaryBarcode || product.sku,
    extraBarcodes,
    brand: product.brand ?? '',
    model: product.model ?? '',
    unit: product.unit ?? 'шт',
    storageLocation: product.storageLocation ?? '',
    min: product.min ?? 1,
    stock: batches.reduce((sum, batch) => sum + batch.qtyAvailable, 0),
    batches,
  };
}

function findProductByLookup(products: Product[], query: string) {
  return products.find((product) => productMatchesLookup(product, query));
}

function findSimilarProducts(products: Product[], input: { name?: string; sku?: string; barcode?: string; brand?: string; model?: string }) {
  const name = normalizeWarehouseKey(input.name ?? '');
  const sku = normalizeWarehouseKey(input.sku ?? '');
  const barcode = normalizeBarcodeValue(input.barcode ?? '');
  const brand = normalizeWarehouseKey(input.brand ?? '');
  const model = normalizeWarehouseKey(input.model ?? '');
  return products.filter((product) => {
    const sameSku = sku && normalizeWarehouseKey(product.sku) === sku;
    const sameBarcode = barcode && productBarcodeList(product).includes(barcode);
    const sameBrandModel = brand && model && normalizeWarehouseKey(product.brand) === brand && normalizeWarehouseKey(product.model) === model;
    const sameName = name && normalizeWarehouseKey(product.name).includes(name);
    return Boolean(sameSku || sameBarcode || sameBrandModel || sameName);
  });
}

function ensureUniqueSku(products: Product[], sku: string, ignoreProductId?: string) {
  const normalizedSku = normalizeWarehouseKey(sku);
  if (!normalizedSku) return false;
  return !products.some((product) => product.id !== ignoreProductId && normalizeWarehouseKey(product.sku) === normalizedSku);
}

function importProductsFromRows(rows: Record<string, unknown>[]) {
  return rows
    .map<Product | null>((row, index) => {
      const name = String(valueFromRow(row, ['название', 'назва', 'товар', 'найменування', 'name']) ?? '').trim();
      const sku = String(valueFromRow(row, ['артикул', 'sku', 'код', 'article']) ?? '').trim();
      const category = String(valueFromRow(row, ['категория', 'категорія', 'category']) ?? '').trim();
      const brand = String(valueFromRow(row, ['бренд', 'brand', 'виробник', 'производитель']) ?? '').trim();
      const model = String(valueFromRow(row, ['модель', 'model']) ?? '').trim();
      const barcode = String(valueFromRow(row, ['штрихкод', 'штрих-код', 'barcode', 'ean']) ?? '').trim();
      const qty = Number(String(valueFromRow(row, ['количество', 'кількість', 'qty', 'quantity']) ?? '0').replace(',', '.')) || 0;
      const price = Number(String(valueFromRow(row, ['цена', 'ціна', 'price', 'cost']) ?? '0').replace(',', '.')) || 0;
      if (!sku) return null;
      const productId = uid(`imp-${index + 1}`);
      const normalizedName = buildNormalizedProductName({ category, brand, model, name });
      const primaryBarcode = normalizeBarcodeValue(barcode || sku || productId);
      return {
        id: productId,
        name: normalizedName || normalizeWarehouseText(name) || sku,
        sku,
        barcode: primaryBarcode,
        extraBarcodes: [] as string[],
        category: category || 'Імпорт',
        brand,
        model,
        unit: 'шт',
        cost: price,
        price,
        stock: qty,
        reserved: 0,
        withEngineer: 0,
        installed: 0,
        min: 1,
        storageLocation: '',
        batches: qty > 0 ? [{
          id: uid('BATCH'),
          productId,
          qtyTotal: qty,
          qtyAvailable: qty,
          purchasePrice: price,
          purchaseDate: today,
          source: 'Імпорт',
          supplier: 'Імпорт',
        }] : [],
      } satisfies Product;
    })
    .filter((product): product is Product => Boolean(product));
}

function mergeImportedProducts(currentProducts: Product[], importedProducts: Product[]) {
  let nextProducts = [...currentProducts];
  importedProducts.forEach((incoming) => {
    const exactMatch = findProductByLookup(nextProducts, incoming.barcode)
      ?? findSimilarProducts(nextProducts, incoming)[0];
    if (exactMatch) {
      const appendedBarcodes = normalizeBarcodeList(...exactMatch.extraBarcodes, incoming.barcode, ...(incoming.extraBarcodes ?? [])).filter((item) => item !== exactMatch.barcode);
      nextProducts = nextProducts.map((product) => (
        product.id !== exactMatch.id
          ? product
          : normalizeStoredProduct({
              ...product,
              name: buildNormalizedProductName({ category: product.category || incoming.category, brand: product.brand || incoming.brand, model: product.model || incoming.model, name: product.name || incoming.name }),
              category: product.category || incoming.category,
              brand: product.brand || incoming.brand,
              model: product.model || incoming.model,
              unit: product.unit || incoming.unit,
              cost: incoming.cost || product.cost,
              price: incoming.price || product.price,
              min: Math.max(product.min, incoming.min),
              storageLocation: product.storageLocation || incoming.storageLocation,
              extraBarcodes: appendedBarcodes,
              batches: [
                ...product.batches,
                ...incoming.batches.map((batch) => ({ ...batch, productId: product.id })),
              ],
            })
      ));
      return;
    }
    nextProducts = [normalizeStoredProduct(incoming), ...nextProducts];
  });
  return nextProducts;
}

function importBankStatementRows(rows: Record<string, unknown>[], mapping: Partial<Record<BankImportField, string>>) {
  return rows
    .map((row) => ({
      id: uid('BANKROW'),
      date: String(resolveMappedValue(row, 'date', mapping) ?? '').trim(),
      amount: parseImportedAmount(resolveMappedValue(row, 'amount', mapping)),
      currency: String(resolveMappedValue(row, 'currency', mapping) ?? '').trim() || 'UAH',
      payer: String(resolveMappedValue(row, 'payer', mapping) ?? '').trim(),
      taxId: String(resolveMappedValue(row, 'taxId', mapping) ?? '').trim(),
      purpose: String(resolveMappedValue(row, 'purpose', mapping) ?? '').trim(),
      documentRef: String(resolveMappedValue(row, 'documentRef', mapping) ?? '').trim(),
      direction: detectBankDirection(resolveMappedValue(row, 'direction', mapping)),
      bankAccount: String(resolveMappedValue(row, 'bankAccount', mapping) ?? '').trim(),
    }))
    .filter((row) => row.amount > 0 || row.payer || row.purpose || row.documentRef);
}

function fifoConsumeBatches(product: Product, qty: number) {
  if (batchTotalAvailable(product) < qty) return null;
  let remaining = qty;
  const allocations: BatchAllocation[] = [];
  const batches = product.batches.map((batch) => {
    if (remaining <= 0 || batch.qtyAvailable <= 0) return batch;
    const taken = Math.min(batch.qtyAvailable, remaining);
    if (taken > 0) {
      allocations.push({ batchId: batch.id, qty: taken, unitCost: batch.purchasePrice, purchaseDate: batch.purchaseDate });
      remaining -= taken;
      return { ...batch, qtyAvailable: Math.max(batch.qtyAvailable - taken, 0) };
    }
    return batch;
  });
  if (remaining > 0) return null;
  const totalCost = allocations.reduce((sum, item) => sum + item.qty * item.unitCost, 0);
  return { batches, allocations, totalCost };
}

function restoreBatchAllocations(product: Product, allocations: BatchAllocation[]) {
  return product.batches.map((batch) => {
    const allocation = allocations.find((item) => item.batchId === batch.id);
    return allocation ? { ...batch, qtyAvailable: batch.qtyAvailable + allocation.qty } : batch;
  });
}

function saleTotals(sale: Sale) {
  const activeItems = sale.items.filter((item) => item.status !== 'Скасовано');
  const sum = activeItems.reduce((total, item) => total + item.price * item.qty, 0);
  const discount = activeItems.reduce((total, item) => total + item.discount, 0);
  const paid = sale.payments.filter(paymentCountsAsApplied).reduce((total, payment) => total + payment.amount, 0);
  const cost = activeItems.filter((item) => item.status === 'Видано').reduce((total, item) => total + item.cost * item.qty, 0);
  const returns = sale.returns.reduce((total, item) => total + item.refund, 0);
  const total = Math.max(sum - discount, 0);
  return { sum, discount, total, paid, debt: Math.max(total - paid, 0), cost, returns, profit: total - cost - returns };
}

function orderTotals(order: ServiceOrder) {
  const works = order.works.reduce((sum, work) => sum + work.price * (work.qty ?? 1), 0);
  const parts = order.parts.reduce((sum, part) => sum + part.price * part.qty, 0);
  const delivery = order.deliveryAmount ?? 0;
  const additionalExpenses = order.additionalExpenses ?? 0;
  const consumables = order.consumablesAmount ?? 0;
  const externalServices = order.externalServicesAmount ?? 0;
  const paid = order.payments.filter(paymentCountsAsApplied).reduce((sum, payment) => sum + payment.amount, 0);
  const purchaseCost = order.parts.filter((part) => part.status !== 'Повернення').reduce((sum, part) => sum + part.cost * part.qty, 0);
  const engineerSalary = order.works.reduce((sum, work) => {
    const engineerName = work.engineer ?? order.engineer;
    const rule = payrollRules.find((item) => item.employee === engineerName);
    return sum + payrollForWork(order, work, rule);
  }, 0);
  const cost = purchaseCost + engineerSalary + delivery + additionalExpenses + consumables + externalServices;
  const total = works + parts + delivery;
  const isFinal = ['Закрито', 'Видано'].includes(order.status);
  return { works, parts, delivery, additionalExpenses, consumables, externalServices, total, paid, debt: Math.max(total - paid, 0), purchaseCost, engineerSalary, cost, profit: total - cost, finalProfit: isFinal ? total - cost : 0, isFinal };
}

function paymentTypeFor(amount: number, debtBeforePayment: number, fallback: PaymentType): PaymentType {
  if (fallback === 'Передплата' || fallback === 'Повернення коштів') return fallback;
  if (amount >= debtBeforePayment) return 'Повна оплата';
  return 'Часткова оплата';
}

function paymentCountsAsApplied(payment: Payment) {
  if (payment.method === 'Картка' && payment.status === 'Проведено') return true;
  return payment.status === 'Підтверджено' || payment.status === 'Скасовано' || payment.status === 'Повернення' || !payment.status;
}

function paymentNeedsConfirmation(payment: Payment) {
  return payment.status === 'Очікує підтвердження' || payment.status === 'Очікує надходження';
}

function paymentProcessingLabel(payment: Payment) {
  if (payment.status === 'Скасовано') return 'Скасовано';
  if (payment.status === 'Повернення') return 'Повернення';
  if (payment.status === 'Проведено') return 'Проведено';
  if (payment.status === 'Очікує надходження') return 'Очікує надходження';
  if (payment.status === 'Очікує підтвердження') return 'Чернетка';
  return 'Підтверджено';
}

function hasUnconfirmedReleasePayments(order: ServiceOrder) {
  return order.payments.some((payment) => (
    payment.amount > 0
    && (
      payment.status === 'Очікує надходження'
      || payment.status === 'Очікує підтвердження'
    )
  ));
}

function orderDebtAmount(order: ServiceOrder) {
  const totalAmount = Math.max(order.repairPrice ?? order.estimatedAmount ?? orderTotals(order).total, 0);
  const paidAmount = order.payments.filter(paymentCountsAsApplied).reduce((sum, payment) => sum + payment.amount, 0);
  return Math.max(totalAmount - paidAmount, 0);
}

function getOrderActStatus(documents: PrintDocument[], orderId: string): ReleaseActStatus {
  const actDocument = documents.find((document) => document.entityType === 'service_order' && document.entityId === orderId && document.kind === 'Акт надання послуг');
  if (!actDocument) return 'Не створено';
  if (actDocument.status === 'Підписано') return 'Підписано';
  if (actDocument.status === 'Роздруковано') return 'Роздруковано';
  return 'Не створено';
}

function canRelease(order: ServiceOrder, actStatus: ReleaseActStatus) {
  return (
    orderDebtAmount(order) === 0
    && order.status === 'Готовий до видачі'
    && actStatus === 'Підписано'
    && !hasUnconfirmedReleasePayments(order)
  );
}

function getBlockReason(order: ServiceOrder, actStatus: ReleaseActStatus): string | null {
  if (orderDebtAmount(order) > 0) return 'Є борг';
  if (hasUnconfirmedReleasePayments(order)) return 'Платіж не підтверджено';
  if (actStatus !== 'Підписано') return 'Акт не підписано';
  if (order.status !== 'Готовий до видачі') return 'Замовлення не завершено';
  return null;
}

function getNextAction(order: ServiceOrder, actStatus: ReleaseActStatus): ActionType {
  if (orderDebtAmount(order) > 0) return 'Оплатити';
  if (hasUnconfirmedReleasePayments(order)) return 'Підтвердити оплату';
  if (actStatus === 'Не створено') return 'Роздрукувати акт';
  if (actStatus === 'Роздруковано') return 'Підписати акт';
  if (canRelease(order, actStatus)) return 'Видати';
  return 'Перейти до ремонту';
}

function getNextActionHint(order: ServiceOrder, actStatus: ReleaseActStatus) {
  const remaining = orderDebtAmount(order);
  const nextAction = getNextAction(order, actStatus);
  if (nextAction === 'Оплатити') return `Залишок до оплати: ${money(remaining)}`;
  if (nextAction === 'Підтвердити оплату') return 'Очікується підтвердження платежу';
  if (nextAction === 'Роздрукувати акт') return 'Акт ще не створено';
  if (nextAction === 'Підписати акт') return 'Акт роздруковано, очікується підпис';
  if (nextAction === 'Видати') return 'Усі умови виконані';
  return 'Замовлення ще не завершено';
}

function findLatestOrderTaxInvoice(taxInvoices: TaxInvoice[], orderId: string) {
  return [...taxInvoices]
    .filter((invoice) => invoice.orderId === orderId)
    .sort((a, b) => (parseDateTime(b.createdAt)?.getTime() ?? 0) - (parseDateTime(a.createdAt)?.getTime() ?? 0))[0];
}

function orderRequiresVatDocument(order: ServiceOrder) {
  return Boolean(order.legalEntity || order.vatStatus);
}

function isOrderDealClosed(order: ServiceOrder, documents: PrintDocument[], taxInvoices: TaxInvoice[]) {
  const actStatus = getOrderActStatus(documents, order.id);
  const taxInvoice = findLatestOrderTaxInvoice(taxInvoices, order.id);
  return (
    orderDebtAmount(order) === 0
    && actStatus === 'Підписано'
    && !hasUnconfirmedReleasePayments(order)
    && (!orderRequiresVatDocument(order) || Boolean(taxInvoice))
  );
}

function orderHasDocumentVersionMismatch(order: ServiceOrder, documents: PrintDocument[]) {
  const currentVersion = order.currentVersion ?? 1;
  return documents.some((document) => (
    document.entityType === 'service_order'
    && document.entityId === order.id
    && ['Рахунок на оплату', 'Акт надання послуг', 'Видаткова накладна', 'Гарантійний талон'].includes(document.kind)
    && (document.orderVersionNo ?? 1) !== currentVersion
  ));
}

function orderDocumentsMismatchText(order: ServiceOrder) {
  return `Дані замовлення ${order.id} змінені після створення документів.`;
}

function pendingPaymentReason(order: ServiceOrder) {
  const blockingPayment = order.payments.find((payment) => payment.amount > 0 && (payment.status === 'Очікує надходження' || payment.status === 'Очікує підтвердження'));
  if (!blockingPayment) return '';
  if (blockingPayment.method === 'Безготівка' && blockingPayment.status === 'Очікує надходження') return 'Оплата по IBAN ще не надійшла';
  if (blockingPayment.status === 'Очікує підтвердження') return 'Платіж не підтверджено';
  return '';
}

function expectedCashAmount(shift: CashShift) {
  return shift.openingCash + shift.cashIncome - shift.cashExpense;
}

function vatFromGross(amount: number, rate = 0.2) {
  const net = amount / (1 + rate);
  return { net, vat: amount - net, gross: amount, rate };
}

function documentVatTotals(amountGross: number) {
  const vat = vatFromGross(amountGross);
  return { amountNet: vat.net, amountVat: vat.vat, amountGross: vat.gross };
}

function invoiceTotals(totalWithVAT: number, vatEnabled = true) {
  if (!vatEnabled) {
    return {
      totalWithoutVAT: totalWithVAT,
      vat: 0,
      totalWithVAT,
    };
  }
  const totalWithoutVAT = totalWithVAT / 1.2;
  const vat = totalWithoutVAT * 0.2;
  return {
    totalWithoutVAT,
    vat,
    totalWithVAT,
  };
}

function numberToWordsUAInteger(value: number, feminine = false): string {
  const unitsMale = ['', 'один', 'два', 'три', 'чотири', "п'ять", 'шість', 'сім', 'вісім', "дев'ять"];
  const unitsFemale = ['', 'одна', 'дві', 'три', 'чотири', "п'ять", 'шість', 'сім', 'вісім', "дев'ять"];
  const teens = ['десять', 'одинадцять', 'дванадцять', 'тринадцять', 'чотирнадцять', "п'ятнадцять", 'шістнадцять', 'сімнадцять', 'вісімнадцять', "дев'ятнадцять"];
  const tens = ['', '', 'двадцять', 'тридцять', 'сорок', "п'ятдесят", 'шістдесят', 'сімдесят', 'вісімдесят', "дев'яносто"];
  const hundreds = ['', 'сто', 'двісті', 'триста', 'чотириста', "п'ятсот", 'шістсот', 'сімсот', 'вісімсот', "дев'ятсот"];
  if (value === 0) return 'нуль';
  const forms = [
    { one: '', few: '', many: '', feminine },
    { one: 'тисяча', few: 'тисячі', many: 'тисяч', feminine: true },
    { one: 'мільйон', few: 'мільйони', many: 'мільйонів', feminine: false },
    { one: 'мільярд', few: 'мільярди', many: 'мільярдів', feminine: false },
  ];
  const plural = (n: number, one: string, few: string, many: string) => {
    const mod10 = n % 10;
    const mod100 = n % 100;
    if (mod10 === 1 && mod100 !== 11) return one;
    if (mod10 >= 2 && mod10 <= 4 && (mod100 < 12 || mod100 > 14)) return few;
    return many;
  };
  const triadToWords = (num: number, triadFeminine: boolean) => {
    const unitWords = triadFeminine ? unitsFemale : unitsMale;
    const parts: string[] = [];
    const h = Math.floor(num / 100);
    const t = Math.floor((num % 100) / 10);
    const u = num % 10;
    if (h) parts.push(hundreds[h]);
    if (t === 1) {
      parts.push(teens[u]);
    } else {
      if (t) parts.push(tens[t]);
      if (u) parts.push(unitWords[u]);
    }
    return parts.join(' ');
  };
  const chunks: string[] = [];
  let rest = value;
  let triadIndex = 0;
  while (rest > 0) {
    const triad = rest % 1000;
    if (triad) {
      const form = forms[triadIndex] ?? forms[forms.length - 1];
      const words = triadToWords(triad, form.feminine);
      const suffix = triadIndex === 0 ? '' : plural(triad, form.one, form.few, form.many);
      chunks.unshift([words, suffix].filter(Boolean).join(' '));
    }
    rest = Math.floor(rest / 1000);
    triadIndex += 1;
  }
  return chunks.join(' ').replace(/\s+/g, ' ').trim();
}

function numberToWordsUA(amount: number) {
  const normalized = Number.isFinite(amount) ? amount : 0;
  let hryvnias = Math.floor(normalized);
  let kop = Math.round((normalized - hryvnias) * 100);
  if (kop === 100) {
    hryvnias += 1;
    kop = 0;
  }
  return `${numberToWordsUAInteger(hryvnias, true).replace(/^./, (char) => char.toUpperCase())} ${hryvnias % 10 === 1 && hryvnias % 100 !== 11 ? 'гривня' : hryvnias % 10 >= 2 && hryvnias % 10 <= 4 && (hryvnias % 100 < 12 || hryvnias % 100 > 14) ? 'гривні' : 'гривень'} ${String(kop).padStart(2, '0')} копійок`;
}

function amountInWordsUA(amount: number) {
  return numberToWordsUA(amount);
}

function renderTemplate(template: string, data: Record<string, string>) {
  return template.replace(/\{\{\s*([^}]+?)\s*\}\}/g, (_, key) => data[key.trim()] ?? '');
}

function nextDocumentNumber(kind: DocumentKind, currentDocuments: PrintDocument[]) {
  const prefix = documentPrefix[kind];
  const lastSequence = currentDocuments
    .filter((document) => document.kind === kind && document.number.startsWith(`${prefix}-`))
    .map((document) => Number(document.number.replace(`${prefix}-`, '')))
    .filter((value) => Number.isFinite(value))
    .reduce((max, value) => Math.max(max, value), 1000);
  return `${prefix}-${lastSequence + 1}`;
}

function clientDetails(name: string, phone?: string): ClientRequisites {
  return clientRequisites[name] ?? { name, edrpou: 'не вказано', address: 'не вказано', contract: 'Без договору, згідно рахунку', contact: phone ?? 'не вказано' };
}

function serviceOrderInvoiceItems(order: ServiceOrder, products: Product[]) {
  const workItems = order.works.map((work) => ({
    label: work.name,
    qty: work.qty ?? 1,
    price: work.price,
    total: work.price * (work.qty ?? 1),
  }));
  const partItems = order.parts
    .filter((part) => part.status !== 'Повернення')
    .map((part) => ({
      label: products.find((item) => item.id === part.productId)?.name ?? part.productId,
      qty: part.qty,
      price: part.price,
      total: part.price * part.qty,
    }));
  const deliveryItems = (order.deliveryAmount ?? 0) > 0
    ? [{
        label: 'Доставка',
        qty: 1,
        price: order.deliveryAmount ?? 0,
        total: order.deliveryAmount ?? 0,
      }]
    : [];
  const items = [...workItems, ...partItems, ...deliveryItems];
  return items.length > 0 ? items : [{
    label: 'Послуги згідно замовлення',
    qty: 1,
    price: 0,
    total: 0,
  }];
}

function serviceOrderInvoiceRows(order: ServiceOrder, products: Product[]) {
  return serviceOrderInvoiceItems(order, products)
    .map((item, index) => `<tr><td class="center">${index + 1}</td><td>${escapeHtml(item.label)}</td><td class="center">${item.qty}</td><td class="right">${escapeHtml(money(item.price))}</td><td class="right">${escapeHtml(money(item.total))}</td></tr>`)
    .join('');
}

function nextTaxInvoiceNumber(currentTaxInvoices: TaxInvoice[]) {
  const lastSequence = currentTaxInvoices
    .map((invoice) => Number((invoice.number ?? '').replace(/\D/g, '')))
    .filter((value) => Number.isFinite(value) && value > 0)
    .reduce((max, value) => Math.max(max, value), 1000);
  return `ПН-${lastSequence + 1}`;
}

function serviceOrderTaxInvoiceLines(order: ServiceOrder, products: Product[]) {
  const serviceLines = order.works.map((work) => ({
    type: 'service' as const,
    label: work.name,
    qty: work.qty ?? 1,
    price: work.price,
    total: work.price * (work.qty ?? 1),
  }));
  const productLines = order.parts
    .filter((part) => part.status !== 'Повернення')
    .map((part) => ({
      type: 'product' as const,
      label: products.find((item) => item.id === part.productId)?.name ?? part.productId,
      qty: part.qty,
      price: part.price,
      total: part.price * part.qty,
    }));
  return [...serviceLines, ...productLines];
}

function buildServiceOrderDocumentSnapshot(
  kind: ServiceOrderDocumentKind,
  order: ServiceOrder,
  products: Product[],
  documentNumber: string,
  company: CompanySettings,
  clientTaxIdValue: string,
  links?: { invoiceNumber?: string; invoiceDate?: string; invoiceId?: string; actId?: string; taxInvoiceId?: string; waybillId?: string; paymentId?: string },
) {
  const totals = orderTotals(order);
  const vatTotals = invoiceTotals(totals.total, company.vatEnabled);
  const serviceItems = order.works.map((work) => ({
    label: work.name,
    qty: work.qty ?? 1,
    price: work.price,
    total: work.price * (work.qty ?? 1),
  }));
  const productItems = order.parts
    .filter((part) => part.status !== 'Повернення')
    .map((part) => ({
      label: products.find((item) => item.id === part.productId)?.name ?? part.productId,
      qty: part.qty,
      price: part.price,
      total: part.price * part.qty,
    }));
  const items = serviceOrderInvoiceItems(order, products);
  const rowsHtml = (rows: typeof items) => rows.length > 0
    ? rows.map((item, index) => `<tr><td class="center">${index + 1}</td><td>${escapeHtml(item.label)}</td><td class="center">${item.qty}</td><td class="right">${escapeHtml(money(item.price))}</td><td class="right">${escapeHtml(money(item.total))}</td></tr>`).join('')
    : '<tr><td class="center">1</td><td>Немає позицій</td><td class="center">0</td><td class="right">0 грн</td><td class="right">0 грн</td></tr>';
  return {
    date: kind === 'Акт надання послуг' ? extractDayKey(order.statusChangedAt) : today,
    orderNumber: documentNumber,
    clientName: order.client,
    phone: order.phone,
    device: order.device,
    serialNumber: order.serial,
    problem: order.issue,
    total: money(totals.total),
    paid: money(totals.paid),
    debt: money(totals.debt),
    managerName: order.manager,
    engineerName: order.engineer,
    clientTaxId: clientTaxIdValue,
    companyName: company.companyName,
    companyEdrpou: company.edrpou,
    companyIban: company.iban,
    companyBank: company.bank,
    companyMfo: company.mfo,
    companyAddress: company.address,
    companyPhone: company.phone,
    companyIpn: company.ipn,
    companyVatCertificate: company.vatCertificate,
    companyVatStatus: company.vatEnabled ? 'Платник ПДВ' : 'Без ПДВ',
    amountNet: money(vatTotals.totalWithoutVAT),
    amountVat: money(vatTotals.vat),
    amountGross: money(vatTotals.totalWithVAT),
    amountWords: numberToWordsUA(vatTotals.totalWithVAT),
    vatWords: numberToWordsUA(vatTotals.vat),
    totalWithVATWords: numberToWordsUA(vatTotals.totalWithVAT),
    vatAmountWords: numberToWordsUA(vatTotals.vat),
    worksRows: order.works.length > 0
      ? order.works.map((work, index) => `<tr><td class="center">${index + 1}</td><td>${escapeHtml(work.name)}</td><td class="center">${work.qty ?? 1}</td><td class="right">${escapeHtml(money(work.price))}</td><td class="right">${escapeHtml(money(work.price * (work.qty ?? 1)))}</td></tr>`).join('')
      : '<tr><td class="center">1</td><td>Роботи згідно замовлення</td><td class="center">1</td><td class="right">0 грн</td><td class="right">0 грн</td></tr>',
    partsRows: order.parts.length > 0
      ? order.parts.map((part, index) => `<tr><td class="center">${index + 1}</td><td>${escapeHtml(products.find((item) => item.id === part.productId)?.name ?? part.productId)}</td><td class="center">${part.qty}</td><td class="right">${escapeHtml(money(part.price))}</td><td class="right">${escapeHtml(money(part.price * part.qty))}</td></tr>`).join('')
      : '<tr><td class="center">1</td><td>Запчастини не використовувались</td><td class="center">0</td><td class="right">0 грн</td><td class="right">0 грн</td></tr>',
    invoiceItemsRows: serviceOrderInvoiceRows(order, products),
    serviceItemsRows: rowsHtml(serviceItems),
    productItemsRows: rowsHtml(productItems),
    invoiceItems: items,
    worksList: order.works.map((work) => `${escapeHtml(work.name)} (${work.qty ?? 1} шт.)`).join('<br />'),
    partsList: order.parts.map((part) => `${escapeHtml(products.find((item) => item.id === part.productId)?.name ?? part.productId)} (${part.qty} шт.)`).join('<br />'),
    invoiceReason: `Замовлення ${order.id} від ${extractDayKey(order.intakeDate)}`,
    linkedInvoiceNumber: links?.invoiceNumber ?? '',
    linkedInvoiceDate: links?.invoiceDate ?? '',
    linkedInvoiceId: links?.invoiceId ?? '',
    linkedActId: links?.actId ?? '',
    linkedTaxInvoiceId: links?.taxInvoiceId ?? '',
    linkedWaybillId: links?.waybillId ?? '',
    linkedPaymentId: links?.paymentId ?? '',
    paymentType: links?.paymentId ? (order.payments.find((payment) => payment.id === links.paymentId)?.method ?? '') : '',
    paymentStatus: links?.paymentId ? (order.payments.find((payment) => payment.id === links.paymentId)?.status ?? '') : '',
  };
}

function repairPaymentStatus(order: ServiceOrder): RepairPaymentStatus {
  const totals = orderTotals(order);
  if (order.payments.some((payment) => payment.type === 'Повернення коштів')) return 'Повернення';
  if (totals.paid <= 0 && totals.debt > 0) return 'Не оплачено';
  if (totals.paid > 0 && totals.debt > 0) return order.payments.some((payment) => payment.type === 'Передплата') ? 'Передплата' : 'Частково оплачено';
  if (totals.debt <= 0 && totals.total > 0) return 'Повністю оплачено';
  return totals.debt > 0 ? 'Борг' : 'Не оплачено';
}

function simpleRepairPaymentStatus(order: ServiceOrder): SimpleRepairPaymentStatus {
  const totals = orderTotals(order);
  if (totals.total <= 0 || totals.paid <= 0) return 'Не оплачено';
  if (totals.debt > 0) return 'Частично оплачено';
  return 'Оплачено';
}

function payrollForWork(order: ServiceOrder, work: ServiceWork, rule?: PayrollRule) {
  const qty = work.qty ?? 1;
  const base = work.price * qty;
  if (work.accrualType === 'fixed_per_unit') return Math.round((work.payrollFixedPerUnit ?? 0) * qty);
  const percent = work.payrollPercent ?? rule?.percent ?? 0;
  return Math.round(base * (percent / 100));
}

function workTypeLabel(work: ServiceWork): SimpleWorkType {
  const serviceType = work.serviceType.toLowerCase();
  const workName = work.name.toLowerCase();
  if (serviceType.includes('регенерац') || workName.includes('регенерац')) return 'Регенерация картриджа';
  if (serviceType.includes('заправ') || workName.includes('заправ')) return 'Заправка картриджа';
  return 'Ремонт техники';
}

function workRateLabel(work: ServiceWork, rule?: PayrollRule) {
  if (work.accrualType === 'fixed_per_unit') return `${money(work.payrollFixedPerUnit ?? 0)} / шт.`;
  return `${work.payrollPercent ?? rule?.percent ?? 0}%`;
}

function orderWorkAmount(work: ServiceWork) {
  return work.price * (work.qty ?? 1);
}

function extractDayKey(value: string) {
  const match = value.match(/(\d{2}\.\d{2}\.\d{4})/);
  return match?.[1] ?? value;
}

function extractMonthKey(value: string) {
  const match = value.match(/\d{2}\.(\d{2}\.\d{4})/);
  return match?.[1] ?? value;
}

function parseDocumentSnapshot(value?: string) {
  if (!value) return null;
  try {
    return JSON.parse(value) as {
      date: string;
      orderNumber: string;
      clientName: string;
      phone: string;
      device: string;
      serialNumber: string;
      problem: string;
      total: string;
      paid: string;
      debt: string;
      managerName: string;
      engineerName: string;
      clientTaxId: string;
      companyName: string;
      companyEdrpou: string;
      companyIban: string;
      companyBank: string;
      companyMfo: string;
      companyAddress: string;
      companyPhone: string;
      companyIpn: string;
      companyVatCertificate: string;
      companyVatStatus: string;
      amountNet: string;
      amountVat: string;
      amountGross: string;
      amountWords: string;
      vatWords: string;
      totalWithVATWords: string;
      vatAmountWords: string;
      worksRows: string;
      partsRows: string;
      invoiceItemsRows: string;
      serviceItemsRows: string;
      productItemsRows: string;
      invoiceReason: string;
      linkedInvoiceNumber: string;
      linkedInvoiceDate: string;
      linkedInvoiceId: string;
      linkedActId: string;
      linkedTaxInvoiceId: string;
      linkedWaybillId: string;
      linkedPaymentId: string;
      paymentType: string;
      paymentStatus: string;
      worksList: string;
      partsList: string;
    };
  } catch {
    return null;
  }
}

function orderProfitBreakdown(order: ServiceOrder) {
  const totals = orderTotals(order);
  const workCost = totals.engineerSalary + (order.consumablesAmount ?? 0) + (order.externalServicesAmount ?? 0);
  const partsCost = totals.purchaseCost;
  return {
    clientMoney: totals.paid,
    companyMoney: Math.max(totals.paid - totals.engineerSalary, 0),
    engineerMoney: totals.engineerSalary,
    workCost,
    partsCost,
    cost: workCost + partsCost + (order.deliveryAmount ?? 0) + (order.additionalExpenses ?? 0),
    profit: totals.total - (workCost + partsCost + (order.deliveryAmount ?? 0) + (order.additionalExpenses ?? 0)),
  };
}

function buildPayrollEmployeeSummaries(orders: ServiceOrder[]) {
  const closedStatuses: OrderStatus[] = ['Видано', 'Закрито'];
  const todayKey = extractDayKey(today);
  const monthKey = extractMonthKey(today);
  const workRows = orders.flatMap((order) =>
    order.works.map((work) => {
      const employee = work.engineer ?? order.engineer;
      const rule = payrollRules.find((item) => item.employee === employee);
      const workAmount = orderWorkAmount(work);
      const paidAndClosed = closedStatuses.includes(order.status) && orderTotals(order).debt <= 0 && orderTotals(order).paid > 0;
      const accrued = paidAndClosed ? payrollForWork(order, work, rule) : 0;
      const workType = workTypeLabel(work);
      return {
        orderId: order.id,
        employee,
        qty: work.qty ?? 1,
        workAmount,
        accrued,
        workType,
        completedAt: paidAndClosed ? order.statusChangedAt : undefined,
      };
    }),
  );

  return payrollRules
    .filter((rule) => rule.role === 'Інженер')
    .map((rule) => {
      const employeeRows = workRows.filter((row) => row.employee === rule.employee);
      const completedRows = employeeRows.filter((row) => row.completedAt);
      const repairAccrued = completedRows.filter((row) => row.workType === 'Ремонт техники').reduce((sum, row) => sum + row.accrued, 0);
      const cartridgeAccrued = completedRows.filter((row) => row.workType !== 'Ремонт техники').reduce((sum, row) => sum + row.accrued, 0);
      return {
        rule,
        ordersCount: new Set(employeeRows.map((row) => row.orderId)).size,
        operations: employeeRows.length,
        accrued: completedRows.reduce((sum, row) => sum + row.accrued, 0),
        dayAccrued: completedRows.filter((row) => row.completedAt && extractDayKey(row.completedAt) === todayKey).reduce((sum, row) => sum + row.accrued, 0),
        monthAccrued: completedRows.filter((row) => row.completedAt && extractMonthKey(row.completedAt) === monthKey).reduce((sum, row) => sum + row.accrued, 0),
        dayOperations: completedRows.filter((row) => row.completedAt && extractDayKey(row.completedAt) === todayKey).length,
        monthOperations: completedRows.filter((row) => row.completedAt && extractMonthKey(row.completedAt) === monthKey).length,
        workAmount: completedRows.reduce((sum, row) => sum + row.workAmount, 0),
        repairAccrued,
        cartridgeAccrued,
        avgOrderAccrued: completedRows.length ? Math.round(completedRows.reduce((sum, row) => sum + row.accrued, 0) / new Set(completedRows.map((row) => row.orderId)).size) : 0,
      };
    });
}

function buildOrderControlState(order: ServiceOrder, allOrders: ServiceOrder[], notifications: ClientNotification[], units: OrderUnit[]) {
  const totals = orderTotals(order);
  const critical: OrderControlIssue[] = [];
  const warnings: OrderControlIssue[] = [];
  const normalizedPhone = normalizePhone(order.phone);
  const relatedUnits = units.filter((unit) => unit.orderId === order.id);
  const duplicateOrders = allOrders.filter((item) => item.id !== order.id && normalizePhone(item.phone) === normalizedPhone && item.device.trim().toLowerCase() === order.device.trim().toLowerCase() && !['Закрито', 'Скасовано'].includes(item.status));
  const hasSentNotification = notifications.some((item) => item.status === 'Відправлено');
  const missingStorageCell = relatedUnits.some((unit) => unit.status === 'На полці' && !unit.locationCode) || (
    ['Готовий до видачі', 'Очікує оплати', 'Очікує клієнта'].includes(order.status)
    && relatedUnits.length > 0
    && !relatedUnits.some((unit) => unit.locationCode)
  );

  if (!order.phone.trim()) critical.push({ code: 'missing_phone', message: 'Не указан телефон клиента.' });
  if (!order.device.trim()) critical.push({ code: 'missing_device', message: 'Не указано устройство.' });
  if (!order.issue.trim()) critical.push({ code: 'missing_issue', message: 'Не указана проблема клиента.' });
  if (order.pendingExtraApproval) critical.push({ code: 'pending_extra_approval', message: 'Є несогласовані додаткові роботи.' });

  if (totals.debt > 0) warnings.push({ code: 'payment_debt', message: `Есть долг клиента: ${money(totals.debt)}.` });
  if (!hasSentNotification) warnings.push({ code: 'client_not_notified', message: 'Клиент ещё не был уведомлён.' });
  if (daysSince(order.statusChangedAt) > 2) warnings.push({ code: 'stale_order', message: `Заказ давно не менялся: ${daysSince(order.statusChangedAt)} дн.` });
  if (missingStorageCell) warnings.push({ code: 'missing_storage_cell', message: 'Не указана ячейка хранения.' });
  if (duplicateOrders.length > 0) warnings.push({ code: 'possible_duplicate', message: `Возможный дубликат заказа: ${duplicateOrders[0].id}.` });

  return { critical, warnings };
}

const simpleNotificationOptions: Array<{ value: NotificationEvent; label: string }> = [
  { value: 'Прийом пристрою', label: 'Принят заказ' },
  { value: 'Діагностика розпочата', label: 'Диагностика начата' },
  { value: 'Діагностика завершена', label: 'Диагностика завершена' },
  { value: 'Потрібне погодження додаткових робіт', label: 'Согласование доп. работ' },
  { value: 'Погодження прийнято', label: 'Согласование принято' },
  { value: 'Погодження відхилено', label: 'Согласование отклонено' },
  { value: 'Очікування запчастини', label: 'Ожидаем запчасть' },
  { value: 'Ремонт розпочато', label: 'Ремонт начат' },
  { value: 'Ремонт завершено', label: 'Ремонт завершён' },
  { value: 'Готово до видачі', label: 'Готово' },
  { value: 'Очікує оплату', label: 'Ожидает оплату' },
  { value: 'Замовлення затримується', label: 'Заказ задерживается' },
  { value: 'Видача', label: 'Выдан' },
];

function notificationDisplay(event: NotificationEvent) {
  const option = simpleNotificationOptions.find((item) => item.value === event);
  return option?.label ?? event;
}

function notificationTemplateForEvent(event: NotificationEvent) {
  return clientNotificationTemplates.find((item) => item.event === event && item.enabled);
}

function renderNotificationText(order: ServiceOrder, event: NotificationEvent, days = 3, extra?: { extraWorkDescription?: string; extraWorkAmount?: number; comment?: string }) {
  const template = notificationTemplateForEvent(event);
  if (!template) return `${order.client}, статус заказа ${order.id}: ${notificationDisplay(event)}.`;
  const totals = orderTotals(order);
  return template.text
    .split('{client}').join(order.client)
    .split('{order}').join(order.id)
    .split('{amount}').join(money(totals.total))
    .split('{extraWorkAmount}').join(extra?.extraWorkAmount ? money(extra.extraWorkAmount) : money(0))
    .split('{extraWorkDescription}').join(extra?.extraWorkDescription ?? 'додаткові роботи уточнюються')
    .split('{days}').join(String(days))
    .split('{date}').join(order.promisedDate)
    .split('{device}').join(order.device)
    .split('{problem}').join(order.issue)
    .split('{status}').join(statusDisplay(order.status))
    .split('{servicePhone}').join(companyRequisites.phone)
    .split('{serviceHours}').join('18:00')
    .split('{serviceAddress}').join(companyRequisites.address);
}

function escapeHtml(value: string) {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

const serviceDocumentPreviewCss = `
  @page { size: A4; margin: 12mm; }
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; background: #fff; color: #111; font-family: "Times New Roman", Times, serif; font-size: 12pt; line-height: 1.25; }
  body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  .sheet { width: 210mm; min-height: 297mm; margin: 0 auto; padding: 8mm 10mm 10mm; background: #fff; }
  .sheet.compact { padding-top: 6mm; }
  .header { display: grid; grid-template-columns: 28mm 1fr; gap: 6mm; align-items: start; margin-bottom: 6mm; }
  .brand-mark { width: 26mm; height: 23mm; border: 1px solid #444; position: relative; display: flex; align-items: center; justify-content: center; font-weight: 700; font-size: 13pt; letter-spacing: 0.5px; }
  .brand-mark::before { content: "СПЕКТР"; position: absolute; top: 1.5mm; left: 2mm; font-size: 6pt; letter-spacing: 0.2px; }
  .brand-mark::after { content: "АС"; position: absolute; bottom: 2mm; right: 2mm; font-size: 16pt; font-weight: 700; }
  .company-block { font-size: 10.5pt; }
  .company-name { font-weight: 700; }
  .small { font-size: 10pt; }
  .tiny { font-size: 9pt; }
  .title { text-align: center; font-weight: 700; font-size: 18pt; margin: 8mm 0 2mm; }
  .subtitle { text-align: center; font-size: 15pt; margin: 0 0 7mm; }
  .doc-line { font-size: 15pt; font-weight: 700; margin: 3mm 0 5mm; }
  .rule { border-top: 1px solid #222; margin: 2mm 0; }
  .meta-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 2mm 8mm; margin-bottom: 4mm; }
  .meta-row { display: grid; grid-template-columns: 34mm 1fr; align-items: end; gap: 2mm; min-height: 6mm; }
  .meta-row .label { white-space: nowrap; }
  .value-line { border-bottom: 1px solid #222; min-height: 5mm; padding: 0 1mm 0.6mm; }
  .doc-table { width: 100%; border-collapse: collapse; margin: 3mm 0 5mm; }
  .doc-table th, .doc-table td { border: 1px solid #333; padding: 1.2mm 1.5mm; vertical-align: top; }
  .doc-table th { font-weight: 700; text-align: center; }
  .right { text-align: right; }
  .center { text-align: center; }
  .justify { text-align: justify; }
  .terms-box { border: 1px solid #333; padding: 2.5mm 3mm; margin-top: 4mm; }
  .terms-box p { margin: 0 0 2mm; text-align: justify; }
  .signature-row { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 6mm; margin-top: 8mm; }
  .signature-row.two { grid-template-columns: 1fr 1fr; }
  .signature-block { display: flex; flex-direction: column; gap: 1mm; }
  .signature-line { border-bottom: 1px solid #222; min-height: 6mm; }
  .totals { width: 92mm; margin-left: auto; margin-top: 2mm; }
  .totals-row { display: grid; grid-template-columns: 1fr 30mm; gap: 4mm; margin-bottom: 1mm; font-weight: 700; }
  .amount-text { margin-top: 4mm; }
  .stamp-note { margin-top: 8mm; }
  .list { margin: 2mm 0 0 5mm; }
  .list li { margin-bottom: 1mm; }
  .spacer-xxl { height: 14mm; }
  .spacer-xl { height: 10mm; }
  .spacer-lg { height: 6mm; }
  .spacer-sm { height: 3mm; }
  .details-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 2mm 8mm; }
`;

const ENGINEER_OVERDUE_HOURS = 24;

const orderStatusFlow: Partial<Record<OrderStatus, OrderStatus[]>> = {
  Прийнято: ['На діагностиці', 'В ремонті', 'Скасовано'],
  'На діагностиці': ['Очікує погодження', 'Не підлягає ремонту', 'Пауза / відкладено'],
  'Очікує погодження': ['Погоджено', 'Скасовано', 'Очікує клієнта'],
  Погоджено: ['Очікує запчастину', 'В ремонті', 'Скасовано'],
  'Очікує запчастину': ['В ремонті', 'Пауза / відкладено', 'Скасовано'],
  'В ремонті': ['На тестуванні', 'Пауза / відкладено', 'Не підлягає ремонту'],
  'На тестуванні': ['Готовий до видачі', 'В ремонті'],
  'Готовий до видачі': ['Очікує оплати', 'Очікує клієнта', 'Видано'],
  'Очікує оплати': ['Видано', 'Очікує клієнта'],
  'Очікує клієнта': ['Видано', 'Скасовано'],
  'Не підлягає ремонту': ['Видано'],
  Видано: ['Закрито'],
  'Пауза / відкладено': ['На діагностиці', 'В ремонті', 'Очікує погодження', 'Скасовано'],
};

const managerStatuses: OrderStatus[] = ['Прийнято', 'Очікує погодження', 'Погоджено', 'Готовий до видачі', 'Очікує оплати', 'Очікує клієнта', 'Видано', 'Скасовано', 'Не підлягає ремонту'];
const engineerStatuses: OrderStatus[] = ['На діагностиці', 'Очікує погодження', 'Очікує запчастину', 'В ремонті', 'На тестуванні', 'Готовий до видачі', 'Не підлягає ремонту', 'Пауза / відкладено'];
const statusDisplayMap: Partial<Record<OrderStatus, string>> = {
  Прийнято: 'Принято',
  'На діагностиці': 'Диагностика',
  'Очікує погодження': 'Ожидает согласование',
  Погоджено: 'Согласовано',
  'Очікує запчастину': 'Ожидает запчасть',
  'В ремонті': 'В ремонте',
  'На тестуванні': 'На тесте',
  'Готовий до видачі': 'Готово к выдаче',
  'Очікує оплати': 'Ожидает оплату',
  'Очікує клієнта': 'Ожидает клиента',
  Видано: 'Выдано',
  Закрито: 'Закрыто',
  Скасовано: 'Отменено',
  'Не підлягає ремонту': 'Не подлежит ремонту',
  'Пауза / відкладено': 'Пауза',
};
const roleDisplayMap: Record<Role, string> = {
  Руководитель: 'Руководитель',
  Адміністратор: 'Администратор',
  Менеджер: 'Менеджер',
  Інженер: 'Инженер',
  Комірник: 'Кладовщик',
  Бухгалтер: 'Бухгалтер',
  Закупник: 'Закупщик',
};
const orderLifecycleDisplay: Array<{ key: OrderStatus; label: string }> = [
  { key: 'Прийнято', label: 'Принято' },
  { key: 'На діагностиці', label: 'Диагностика' },
  { key: 'Очікує погодження', label: 'Согласование' },
  { key: 'В ремонті', label: 'Ремонт' },
  { key: 'Очікує запчастину', label: 'Ожидает запчасть' },
  { key: 'На тестуванні', label: 'Тест' },
  { key: 'Готовий до видачі', label: 'Готово к выдаче' },
  { key: 'Видано', label: 'Выдано' },
];
const engineerLifecycleDisplay: Array<{ key: OrderStatus; label: string }> = [
  { key: 'На діагностиці', label: 'Диагностика' },
  { key: 'В ремонті', label: 'Ремонт' },
  { key: 'Очікує запчастину', label: 'Ожидает запчасть' },
  { key: 'На тестуванні', label: 'Тест' },
  { key: 'Готовий до видачі', label: 'Готово к выдаче' },
];
const managerLifecycleDisplay: Array<{ key: OrderStatus; label: string }> = [
  { key: 'Прийнято', label: 'Принято' },
  { key: 'Очікує погодження', label: 'Согласование' },
  { key: 'Готовий до видачі', label: 'Готово к выдаче' },
  { key: 'Очікує оплати', label: 'Ожидает оплату' },
  { key: 'Очікує клієнта', label: 'Ожидает клиента' },
  { key: 'Видано', label: 'Выдано' },
];
const statusStageMap: Partial<Record<OrderStatus, string>> = {
  Прийнято: 'Принято',
  'На діагностиці': 'Диагностика',
  'Очікує погодження': 'Согласование',
  Погоджено: 'Согласовано',
  'В ремонті': 'Ремонт',
  'Очікує запчастину': 'Ожидает запчасть',
  'На тестуванні': 'Тест',
  'Готовий до видачі': 'Готово к выдаче',
  'Очікує оплати': 'Ожидает оплату',
  'Очікує клієнта': 'Ожидает клиента',
  'Не підлягає ремонту': 'Не подлежит ремонту',
  Видано: 'Выдано',
  Закрито: 'Закрыто',
  Скасовано: 'Отменено',
  'Пауза / відкладено': 'Пауза',
};

function statusDisplay(status: OrderStatus) {
  return statusDisplayMap[status] ?? status;
}

function statusStage(status: OrderStatus) {
  return statusStageMap[status] ?? statusDisplay(status);
}

function transitionDisplay(status: OrderStatus) {
  const map: Partial<Record<OrderStatus, string>> = {
    'На діагностиці': 'Перевести в диагностику',
    'Очікує погодження': 'Отправить на согласование',
    Погоджено: 'Подтвердить согласование',
    'В ремонті': 'Перевести в ремонт',
    'Очікує запчастину': 'Перевести в ожидание запчасти',
    'На тестуванні': 'Передать на тест',
    'Готовий до видачі': 'Перевести в готово к выдаче',
    'Очікує оплати': 'Перевести в ожидание оплаты',
    'Очікує клієнта': 'Перевести в ожидание клиента',
    Видано: 'Выдано',
    'Не підлягає ремонту': 'Не подлежит ремонту',
    Скасовано: 'Отменено',
  };
  return map[status] ?? statusDisplay(status);
}

function roleDisplay(role: Role) {
  return roleDisplayMap[role] ?? role;
}

function roleWorkspaceHint(role: Role) {
  if (role === 'Менеджер') return 'приём заказов, клиент, оплата, выдача';
  if (role === 'Інженер') return 'ремонт, запчасти, тестирование';
  if (role === 'Комірник') return 'склад, прихід, рух, залишки';
  if (role === 'Бухгалтер') return 'каса, документи, вигрузка в бухгалтерію';
  if (role === 'Адміністратор') return 'доступи, співробітники, повний контроль системи';
  if (role === 'Руководитель') return 'контроль заказов, сотрудников и финансов';
  return 'рабочее пространство CRM';
}

function authModeLabel(mode: AuthMode) {
  return mode === 'phone_code' ? 'Телефон + код' : 'Email + пароль';
}

function normalizePhone(value: string) {
  return value.replace(/\D/g, '');
}

function normalizeLooseText(value: string) {
  return value.trim().toLowerCase().replace(/\s+/g, ' ');
}

function normalizeOrderSearchCode(value: string) {
  return value
    .trim()
    .toLowerCase()
    .replace(/[\s\-–—_/]/g, '')
    .replace(/зн/g, 'zn');
}

function extractDigits(value: string) {
  return value.replace(/\D/g, '');
}

function matchesOrderSearch(order: ServiceOrder, rawNeedle: string) {
  const needle = normalizeLooseText(rawNeedle);
  if (!needle) return true;
  const normalizedCodeNeedle = normalizeOrderSearchCode(rawNeedle);
  const orderCode = normalizeOrderSearchCode(order.id);
  const orderNumber = extractDigits(order.id);
  const needleDigits = extractDigits(rawNeedle);

  if (normalizedCodeNeedle && (orderCode.includes(normalizedCodeNeedle) || normalizedCodeNeedle.includes(orderCode))) return true;
  if (needleDigits && (orderNumber.includes(needleDigits) || needleDigits === orderNumber)) return true;

  const phoneDigits = extractDigits(order.phone);
  if (needleDigits && phoneDigits.includes(needleDigits)) return true;

  return [
    order.client,
    order.phone,
    order.device,
    order.serial,
    order.issue,
    order.locationCode ?? '',
  ].some((value) => normalizeLooseText(String(value)).includes(needle));
}

function findExactOrderMatch(orders: ServiceOrder[], rawNeedle: string) {
  const normalizedCodeNeedle = normalizeOrderSearchCode(rawNeedle);
  const needleDigits = extractDigits(rawNeedle);
  return orders.find((order) => {
    const orderCode = normalizeOrderSearchCode(order.id);
    const orderNumber = extractDigits(order.id);
    if (normalizedCodeNeedle && orderCode === normalizedCodeNeedle) return true;
    if (needleDigits && orderNumber === needleDigits) return true;
    return false;
  });
}

function matchesClientSearch(client: ClientRecord, rawNeedle: string) {
  const needle = normalizeLooseText(rawNeedle);
  if (!needle) return true;
  const needleDigits = extractDigits(rawNeedle);
  const phoneDigits = extractDigits(client.phone);
  const taxDigits = extractDigits(client.taxId ?? '');
  if (needleDigits && (phoneDigits.includes(needleDigits) || taxDigits.includes(needleDigits))) return true;
  return [
    client.name,
    client.phone,
    client.email ?? '',
    client.taxId ?? '',
  ].some((value) => normalizeLooseText(String(value)).includes(needle));
}

function findClientBySearch(clients: ClientRecord[], rawNeedle: string) {
  const needle = rawNeedle.trim();
  if (!needle) return null;
  const exactPhoneDigits = extractDigits(rawNeedle);
  const exactName = normalizeLooseText(rawNeedle);
  return clients.find((client) => {
    if (exactPhoneDigits && extractDigits(client.phone) === exactPhoneDigits) return true;
    return normalizeLooseText(client.name) === exactName || normalizeLooseText(client.taxId ?? '') === exactName;
  }) ?? clients.find((client) => matchesClientSearch(client, rawNeedle)) ?? null;
}

function findExactClientMatch(clients: ClientRecord[], rawNeedle: string) {
  const needle = rawNeedle.trim();
  if (!needle) return null;
  const exactPhoneDigits = extractDigits(rawNeedle);
  const exactName = normalizeLooseText(rawNeedle);
  return clients.find((client) => {
    if (exactPhoneDigits && extractDigits(client.phone) === exactPhoneDigits) return true;
    return normalizeLooseText(client.name) === exactName || normalizeLooseText(client.taxId ?? '') === exactName;
  }) ?? null;
}

function defaultPermissionsForRole(role: Role): UserPermissions {
  if (role === 'Адміністратор') {
    return { canAccessWarehouse: true, canAccessFinance: true, canAccessEmployees: true, canAccessSettings: true, canAccessReports: true };
  }
  if (role === 'Руководитель') {
    return { canAccessWarehouse: true, canAccessFinance: true, canAccessEmployees: true, canAccessSettings: false, canAccessReports: true };
  }
  if (role === 'Менеджер') {
    return { canAccessWarehouse: false, canAccessFinance: false, canAccessEmployees: false, canAccessSettings: false, canAccessReports: false };
  }
  if (role === 'Інженер') {
    return { canAccessWarehouse: false, canAccessFinance: false, canAccessEmployees: false, canAccessSettings: false, canAccessReports: false };
  }
  if (role === 'Бухгалтер') {
    return { canAccessWarehouse: false, canAccessFinance: true, canAccessEmployees: false, canAccessSettings: false, canAccessReports: true };
  }
  if (role === 'Комірник') {
    return { canAccessWarehouse: true, canAccessFinance: false, canAccessEmployees: false, canAccessSettings: false, canAccessReports: false };
  }
  if (role === 'Закупник') {
    return { canAccessWarehouse: true, canAccessFinance: false, canAccessEmployees: false, canAccessSettings: false, canAccessReports: true };
  }
  return { canAccessWarehouse: false, canAccessFinance: false, canAccessEmployees: false, canAccessSettings: false, canAccessReports: false };
}

function normalizeUserRecord(user: User): User {
  const defaultPermissions = defaultPermissionsForRole(user.role);
  return {
    ...user,
    phone: user.phone ? normalizePhone(String(user.phone)) : undefined,
    permissions: {
      ...defaultPermissions,
      ...user.permissions,
    },
  };
}

function normalizeUsers(records: User[]) {
  return records.map(normalizeUserRecord);
}

function hasAccess(role: Role, page: Page) {
  return rolePageAccess[role].includes(page);
}

function hasPermissionBasedPageAccess(user: User, page: Page) {
  if (user.role === 'Адміністратор') return true;
  if (['parts', 'purchases', 'storage', 'movements'].includes(page)) return user.permissions.canAccessWarehouse;
  if (['finance', 'cash', 'documents', 'bank-import', 'tax-invoices'].includes(page)) return user.permissions.canAccessFinance;
  if (['employee-control', 'team', 'payroll'].includes(page)) return user.permissions.canAccessEmployees;
  if (page === 'reports') return user.permissions.canAccessReports;
  if (page === 'settings') return user.permissions.canAccessSettings;
  return false;
}

function canRoleViewPage(user: User, targetPage: Page) {
  return hasAccess(user.role, targetPage) || hasPermissionBasedPageAccess(user, targetPage);
}

function hasPermissionOverride(user: User, permission: Permission) {
  if (user.role === 'Адміністратор') return true;
  if (permission.startsWith('stock.') || permission.startsWith('purchases.')) return user.permissions.canAccessWarehouse;
  if (permission.startsWith('finance.')) return user.permissions.canAccessFinance;
  if (permission.startsWith('reports.')) return user.permissions.canAccessReports;
  if (permission.startsWith('settings.')) return user.permissions.canAccessSettings;
  return false;
}

function getManagerUserId(order: ServiceOrder, users: User[]) {
  return users.find((user) => user.role === 'Менеджер' && user.name === order.manager)?.id ?? '';
}

function getEngineerAssignedUserId(order: ServiceOrder, users: User[]) {
  return order.assignedTo ?? users.find((user) => user.role === 'Інженер' && user.name === order.engineer)?.id ?? '';
}

function getOrderCreatorName(order: ServiceOrder, users: User[]) {
  if (!order.createdByUserId) return 'не указано';
  return users.find((user) => user.id === order.createdByUserId)?.name ?? 'не указано';
}

function getInternalMessageAuthorRole(message: InternalMessage, users: User[]) {
  return users.find((user) => user.name === message.createdBy)?.role ?? 'Система';
}

function navLabelForRole(page: Page, role: Role) {
  if (role === 'Менеджер') {
    if (page === 'dashboard') return 'Головна';
    if (page === 'my-orders') return 'Мої замовлення';
    if (page === 'clients') return 'Клієнти';
    if (page === 'cash') return 'Оплати / Каса';
    if (page === 'problem-clients') return 'Контроль';
  }
  if (role === 'Інженер') {
    if (page === 'dashboard') return 'Рабочий стол';
    if (page === 'my-orders') return 'Мои ремонты';
  }
  if (role === 'Руководитель') {
    if (page === 'dashboard') return 'Контроль';
    if (page === 'orders') return 'Все заказы';
    if (page === 'finance') return 'Финансы';
    if (page === 'reports') return 'Отчёты';
    if (page === 'team') return 'Сотрудники';
  }
  return navItems.find((item) => item.id === page)?.label ?? page;
}

function roleCanSetOrderStatus(role: Role, status: OrderStatus) {
  if (role === 'Адміністратор') return true;
  if (role === 'Менеджер') return managerStatuses.includes(status);
  if (role === 'Інженер') return engineerStatuses.includes(status);
  return false;
}

function daysSince(dateText: string) {
  const date = parseDateTime(dateText);
  if (!date) return 0;
  const now = new Date(2026, 3, 21, 18, 0);
  return Math.max(0, Math.floor((now.getTime() - date.getTime()) / 86400000));
}

function parseDateTime(dateText: string) {
  const match = dateText.match(/(\d{2})\.(\d{2})\.(\d{4})(?:\s+(\d{2}):(\d{2}))?/);
  if (!match) return null;
  const [, day, month, year, hours = '00', minutes = '00'] = match;
  return new Date(Number(year), Number(month) - 1, Number(day), Number(hours), Number(minutes));
}

function isDateInPeriod(dateText: string, period: AccountingPeriod) {
  const date = parseDateTime(dateText);
  const current = parseDateTime(today);
  if (!date || !current) return false;
  if (period === 'today') return extractDayKey(dateText) === extractDayKey(today);
  const diffDays = Math.floor((current.getTime() - date.getTime()) / 86400000);
  if (period === 'week') return diffDays >= 0 && diffDays < 7;
  if (period === 'month') return date.getMonth() === current.getMonth() && date.getFullYear() === current.getFullYear();
  return false;
}

function hoursSince(dateText: string) {
  const date = parseDateTime(dateText);
  if (!date) return 0;
  const now = new Date(2026, 3, 21, 18, 0);
  return Math.max(0, Math.floor((now.getTime() - date.getTime()) / 3600000));
}

function minutesSince(dateText: string) {
  const date = parseDateTime(dateText);
  if (!date) return 0;
  const now = new Date(2026, 3, 21, 18, 0);
  return Math.max(0, Math.floor((now.getTime() - date.getTime()) / 60000));
}

function engineerElapsedLabel(dateText?: string) {
  if (!dateText) return '0хв';
  const totalMinutes = minutesSince(dateText);
  const days = Math.floor(totalMinutes / 1440);
  const hours = Math.floor((totalMinutes % 1440) / 60);
  const minutes = totalMinutes % 60;
  if (days > 0) return `${days}д ${hours}г`;
  if (hours > 0) return `${hours}г ${minutes}хв`;
  return `${minutes}хв`;
}

function buildServiceProblems(orders: ServiceOrder[], documents: PrintDocument[]) {
  const problems: ServiceProblem[] = [];
  const hasDocument = (orderId: string, kinds: DocumentKind[]) => documents.some((document) => document.entityType === 'service_order' && document.entityId === orderId && kinds.includes(document.kind));
  const addProblem = (order: ServiceOrder, type: ServiceProblemType, description: string, level: ServiceProblemLevel, days: number, target: ServiceProblem['target'], responsible = order.manager) => {
    problems.push({
      id: `${order.id}-${type}`,
      orderId: order.id,
      client: order.client,
      type,
      description,
      days,
      responsible,
      engineer: order.engineer,
      manager: order.manager,
      level,
      target,
    });
  };

  orders.forEach((order) => {
    const totals = orderTotals(order);
    const statusAge = daysSince(order.statusChangedAt);
    const actAge = order.actIssuedAt ? daysSince(order.actIssuedAt) : statusAge;
    const hasAct = hasDocument(order.id, ['Акт надання послуг', 'Акт виконаних робіт', 'Акт видачі']);
    const hasDelivery = hasDocument(order.id, ['Видаткова накладна']);
    const hasTaxEvent = order.vatStatus === 'Очікує ПН' || order.vatStatus === 'ПН створено' || order.vatStatus === 'ПН зареєстрована' || order.vatStatus === 'Прострочена реєстрація ПН';
    const hasActOrPayment = hasAct || totals.paid > 0;
    const isClosed = ['Закрито', 'Видано', 'Скасовано'].includes(order.status);

    if (['Готовий до видачі', 'Очікує клієнта', 'Очікує оплати'].includes(order.status) && statusAge > 1 && !isClosed) {
      addProblem(order, 'Заказ готовий, але не виданий', 'Заказ готовий більше 24 годин, але клієнту ще не виданий.', 'Важливо', statusAge, 'order', order.manager);
    }
    if (hasAct && totals.paid === 0 && actAge > 1) {
      addProblem(order, 'Акт є, оплата відсутня', 'Акт сформовано, але немає жодної підтвердженої оплати.', 'Критично', actAge, 'payment', order.manager);
    }
    if (order.parts.length > 0 && !hasDelivery) {
      addProblem(order, 'Запчастини списані, але немає видаткової накладної', 'У заказі є складські позиції, але видаткова накладна не створена.', 'Критично', statusAge, 'document', order.manager);
    }
    if (['Готовий до видачі', 'Очікує оплати', 'Очікує клієнта'].includes(order.status) && !hasAct) {
      addProblem(order, 'Немає акта виконаних робіт', 'Робота завершена, але акт ще не сформовано.', 'Критично', statusAge, 'act', order.manager);
    }
    if (hasActOrPayment && !hasTaxEvent) {
      addProblem(order, 'Немає підстави для податкової накладної', 'Є акт або оплата, але заказ не позначений як готовий до НН.', 'Критично', Math.max(actAge, daysSince(order.debtSince ?? order.intakeDate)), 'act', 'Бухгалтер');
    }
    if (order.status === 'В ремонті' && statusAge > 2) {
      addProblem(order, 'Заказ без руху', 'Заказ у роботі без зміни статусу більше 2 днів.', 'Важливо', statusAge, 'order', order.engineer);
    }
    if (order.actIssuedAt && !order.actReturnedAt && actAge > 3) {
      addProblem(order, 'Акт не повернений клієнтом', 'Акт видано, але підпис клієнта не зафіксовано.', 'Важливо', actAge, 'act', order.manager);
    }
    if (totals.paid > 0 && totals.paid < totals.total && !isClosed) {
      addProblem(order, 'Є заборгованість по заказу', `Оплачено ${money(totals.paid)} із ${money(totals.total)}.`, 'Критично', daysSince(order.debtSince ?? order.intakeDate), 'payment', order.manager);
    }
  });

  return problems.sort((a, b) => {
    const priority = { Критично: 0, Важливо: 1, Інформація: 2 };
    return priority[a.level] - priority[b.level] || b.days - a.days;
  });
}

function defaultPageForUser(user: User): Page {
  if (user.role === 'Адміністратор') return 'dashboard';
  if (user.role === 'Руководитель') return 'dashboard';
  if (user.role === 'Менеджер') return 'dashboard';
  if (user.role === 'Інженер') return 'my-orders';
  if (user.role === 'Бухгалтер') return 'finance';
  if (user.role === 'Комірник' || user.role === 'Закупник') return 'parts';
  return 'dashboard';
}

const adminPreviewRoles: Role[] = ['Адміністратор', 'Руководитель', 'Менеджер', 'Інженер', 'Бухгалтер', 'Комірник', 'Закупник'];

function adminPreviewRoleLabel(role: Role) {
  if (role === 'Руководитель') return 'Керівник';
  return role;
}

function buildPreviewUser(role: Role, currentUser: User) {
  if (role === 'Адміністратор') return currentUser;
  return normalizeUserRecord({
    ...currentUser,
    id: `preview-${role}`,
    name: `${adminPreviewRoleLabel(role)} · preview`,
    role,
    permissions: defaultPermissionsForRole(role),
  });
}

function simpleRepairStatus(status: OrderStatus): 'Прийнято' | 'В ремонті' | 'Готово' | 'Видано' {
  if (status === 'Видано' || status === 'Закрито') return 'Видано';
  if (['Готовий до видачі', 'Очікує оплати', 'Очікує клієнта', 'Не підлягає ремонту', 'Готово', 'Оплачено'].includes(status)) return 'Готово';
  if (status === 'Прийнято') return 'Прийнято';
  return 'В ремонті';
}

function strictManagerWorkflowStatus(order: ServiceOrder) {
  const remaining = clientOrderDebt(order);
  if (order.status === 'Скасовано') return 'Скасовано';
  if (order.status === 'Закрито') return 'Закрито';
  if (order.status === 'Видано') return remaining > 0 ? 'Видано з боргом' : 'Видано';
  if (order.status === 'Готовий до видачі') return 'Готово до видачі';
  if (order.status === 'Очікує оплати') return 'Очікує оплату';
  if (order.engineerWorkCompletedAt || order.returnedToCellAt) return 'Готово до видачі';
  if (order.engineerAcceptedAt || ['На діагностиці', 'В ремонті', 'На тестуванні', 'Очікує запчастину'].includes(order.status)) return 'В роботі';
  if (order.engineer) return 'Інженер призначений';
  return 'Прийнято';
}

function simpleRepairStatusClass(status: ReturnType<typeof simpleRepairStatus>) {
  if (status === 'Прийнято') return 'tag tag-gray';
  if (status === 'В ремонті') return 'tag tag-blue';
  if (status === 'Готово') return 'tag tag-green';
  if (status === 'Видано') return 'tag tag-gray';
  return 'tag tag-blue';
}

function managerOrderStatusLabel(status: OrderStatus) {
  if (status === 'Оплачено') return 'Оплачено';
  return simpleRepairStatus(status);
}

function managerOrderPipelineStage(status: OrderStatus): 'Прийнято' | 'В ремонті' | 'Готово' | 'Видано' {
  if (status === 'Видано' || status === 'Закрито') return 'Видано';
  if (status === 'Оплачено' || ['Готовий до видачі', 'Очікує оплати', 'Очікує клієнта', 'Не підлягає ремонту', 'Готово'].includes(status)) return 'Готово';
  if (status === 'Прийнято') return 'Прийнято';
  return 'В ремонті';
}

function contractOrderAmount(order: ServiceOrder) {
  return order.repairPrice ?? order.estimatedAmount ?? orderTotals(order).total;
}

function contractUsage(contract: ContractRecord, orders: ServiceOrder[]) {
  const contractOrders = orders.filter((order) => order.contractId === contract.id);
  const used = contractOrders.reduce((sum, order) => (
    sum + (order.contractAccountedAt ? contractOrderAmount(order) : 0)
  ), 0);
  return {
    used,
    remaining: Math.max(contract.amount - used, 0),
    contractOrders,
  };
}

function contractClosedAmount(contract: ContractRecord, acts: ContractActRecord[]) {
  return acts
    .filter((act) => act.contractId === contract.id)
    .reduce((sum, act) => sum + act.amount, 0);
}

function actPaymentStatusClass(status: ContractActStatus) {
  if (status === 'Оплачен') return 'tag tag-green';
  if (status === 'Частично оплачен') return 'tag tag-amber';
  return 'tag tag-gray';
}

function recalculateActStatus(act: Pick<ContractActRecord, 'amount' | 'paidAmount'>): ContractActStatus {
  if (act.amount > 0 && act.paidAmount >= act.amount - 0.01) return 'Оплачен';
  if (act.paidAmount > 0) return 'Частично оплачен';
  return 'Не оплачено';
}

function isSameClientOrder(client: ClientRecord, order: ServiceOrder) {
  const clientPhone = client.phone.replace(/\D/g, '');
  const orderPhone = order.phone.replace(/\D/g, '');
  if (clientPhone && orderPhone) return clientPhone === orderPhone;
  return client.name.trim().toLowerCase() === order.client.trim().toLowerCase();
}

function findClientRecord(clients: ClientRecord[], name: string, phone?: string, taxId?: string) {
  const normalizedPhone = normalizeImportedPhone(phone).replace(/\D/g, '');
  const normalizedTaxId = normalizeTaxId(taxId);
  return clients.find((client) => {
    if (normalizedTaxId && normalizeTaxId(client.taxId) === normalizedTaxId) return true;
    if (normalizedPhone && client.phone.replace(/\D/g, '') === normalizedPhone) return true;
    return client.name.trim().toLowerCase() === name.trim().toLowerCase();
  });
}

function clientOrderDebt(order: ServiceOrder) {
  return orderDebtSnapshot(order).remainingDebt;
}

function confirmedOrderPaymentsTotal(order: ServiceOrder) {
  return order.payments
    .filter(paymentCountsAsApplied)
    .reduce((sum, payment) => sum + payment.amount, 0);
}

function orderDebtSnapshot(order: ServiceOrder) {
  const total = Math.max(order.repairPrice ?? order.estimatedAmount ?? orderTotals(order).total, 0);
  const paid = confirmedOrderPaymentsTotal(order);
  const remaining = Math.max(total - paid, 0);
  const baseDebt = order.issuedDebtAmount && order.issuedDebtAmount > 0
    ? Math.max(order.issuedDebtAmount, remaining)
    : remaining;
  const paidTowardDebt = order.issuedDebtAmount && order.issuedDebtAmount > 0
    ? Math.max(baseDebt - remaining, 0)
    : 0;
  const status: OrderDebtLifecycleStatus = remaining <= 0
    ? 'Закрыт'
    : paidTowardDebt > 0
      ? 'В работе (частично оплачен)'
      : 'Новый долг';
  return {
    total,
    paid,
    remainingDebt: remaining,
    debtAmount: Math.max(baseDebt, 0),
    paidTowardDebt,
    status,
  };
}

function clientDebtSummary(client: ClientRecord, orders: ServiceOrder[]) {
  const debtOrders = orders.filter((order) => isSameClientOrder(client, order) && clientOrderDebt(order) > 0);
  return {
    totalDebt: debtOrders.reduce((sum, order) => sum + clientOrderDebt(order), 0),
    debtOrdersCount: debtOrders.length,
  };
}

function clientUncontractedOrders(client: ClientRecord, orders: ServiceOrder[]) {
  return orders.filter((order) => isSameClientOrder(client, order) && !order.contractId);
}

function mapSimpleMethodToPaymentMethod(method: 'наличные' | 'карта' | 'перевод'): PaymentMethod {
  if (method === 'карта') return 'Картка';
  if (method === 'перевод') return 'Безготівка';
  return 'Готівка';
}

function mapSimpleKindToPaymentType(kind: SimpleLedgerPaymentKind): PaymentType {
  if (kind === 'предоплата') return 'Передплата';
  if (kind === 'частичная') return 'Часткова оплата';
  return 'Повна оплата';
}

function simpleLedgerPaymentLabel(order: ServiceOrder) {
  const total = Math.max(order.repairPrice ?? order.estimatedAmount ?? orderTotals(order).total, 0);
  const paid = order.payments.filter(paymentCountsAsApplied).reduce((sum, payment) => sum + payment.amount, 0);
  const remaining = Math.max(total - paid, 0);
  if (total > 0 && remaining <= 0) return 'Оплачено';
  if (paid > 0 && remaining > 0) return 'Частково';
  if ((order.legalEntity || order.contractId) && simpleRepairStatus(order.status) === 'Видано' && remaining > 0) return 'Борг';
  return 'Не оплачено';
}

function inferToastTone(message: string): AppToast['tone'] {
  const normalized = message.toLowerCase();
  if (normalized.includes('не вдалося') || normalized.includes('не можна') || normalized.includes('неможливо') || normalized.includes('помилка') || normalized.includes('недостатньо')) return 'error';
  if (normalized.includes('ризик') || normalized.includes('простроч') || normalized.includes('завис') || normalized.includes('увага') || normalized.includes('без оплати')) return 'warning';
  if (normalized.includes('створено') || normalized.includes('прийнято') || normalized.includes('додано') || normalized.includes('оновлено') || normalized.includes('видано') || normalized.includes('відправлено') || normalized.includes('підтверджено')) return 'success';
  return 'info';
}

function EmployeeLoginScreen({
  phone,
  code,
  challenge,
  error,
  hint,
  onPhoneChange,
  onCodeChange,
  onRequestCode,
  onVerifyCode,
  onBack,
}: {
  phone: string;
  code: string;
  challenge: LoginChallenge | null;
  error: string;
  hint: string;
  onPhoneChange: (value: string) => void;
  onCodeChange: (value: string) => void;
  onRequestCode: () => void;
  onVerifyCode: () => void;
  onBack: () => void;
}) {
  return (
    <div
      className="app-shell"
      style={{
        minHeight: '100vh',
        display: 'grid',
        placeItems: 'center',
        padding: '24px',
        background: 'linear-gradient(180deg, #f8fafc 0%, #eef2ff 100%)',
      }}
    >
      <section
        className="panel"
        style={{
          width: '100%',
          maxWidth: '560px',
          padding: '32px',
          borderRadius: '24px',
          boxShadow: '0 24px 60px rgba(15, 23, 42, 0.12)',
          border: '1px solid rgba(148, 163, 184, 0.18)',
        }}
      >
        <div style={{ display: 'grid', gap: '10px', marginBottom: '24px' }}>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '12px' }}>
            <div>
              <div style={{ fontSize: '13px', fontWeight: 700, letterSpacing: '0.08em', textTransform: 'uppercase', color: '#0f766e' }}>СПЕКТР-АС CRM</div>
              <div style={{ fontSize: '14px', color: '#64748b', marginTop: '4px' }}>Вхід співробітника</div>
            </div>
          </div>
          <div>
            <h1 style={{ margin: 0, fontSize: '34px', lineHeight: 1.05 }}>Телефон + одноразовий код</h1>
            <p style={{ margin: '10px 0 0', color: '#64748b', fontSize: '15px', lineHeight: 1.5 }}>
              Без паролів, без реєстрації, без зайвих кроків. Увійти може тільки співробітник, чий телефон є в базі.
            </p>
          </div>
        </div>
        {!challenge ? (
          <>
            <label>
              Телефон
              <input value={phone} onChange={(event) => onPhoneChange(event.target.value)} placeholder="+380..." style={{ marginTop: '8px' }} />
            </label>
            <div className="action-row" style={{ marginTop: '18px' }}>
              <button type="button" className="submit-button" onClick={onRequestCode} style={{ minWidth: '170px' }}>Отримати код</button>
            </div>
          </>
        ) : (
          <>
            <label>
              Введіть код
              <input value={code} onChange={(event) => onCodeChange(event.target.value)} placeholder="4-6 цифр" style={{ marginTop: '8px' }} />
              <small>Тестовий код для входу: 1111. Телефон: {challenge.phone}.</small>
            </label>
            <div className="action-row" style={{ marginTop: '18px' }}>
              <button type="button" className="submit-button" onClick={onVerifyCode} style={{ minWidth: '170px' }}>Увійти</button>
              <button type="button" onClick={onBack}>Змінити телефон</button>
            </div>
          </>
        )}
        {error && <div className="empty-state" style={{ marginTop: '16px', color: '#b91c1c' }}>{error}</div>}
        {hint && <div className="empty-state" style={{ marginTop: '16px', color: '#166534' }}>{hint}</div>}
      </section>
    </div>
  );
}

export function WarehouseCRM() {
  const loginCodeLifetimeMs = Number((import.meta as ImportMeta & { env?: Record<string, string | undefined> }).env?.VITE_LOGIN_CODE_TTL_MS ?? 180000);
  const [page, setPage] = useState<Page>(() => {
  return (localStorage.getItem('page') as Page) || 'dashboard';
});

useEffect(() => {
  localStorage.setItem('page', page);
}, [page]);
  const [sessionUserId, setSessionUserId] = useState(() => localStorage.getItem('sessionUserId') || '');
  const [activeUserId, setActiveUserId] = useState(() => localStorage.getItem('activeUserId') || '');
useEffect(() => {
  if (sessionUserId) localStorage.setItem('sessionUserId', sessionUserId);
  else localStorage.removeItem('sessionUserId');
}, [sessionUserId]);
useEffect(() => {
  if (activeUserId) localStorage.setItem('activeUserId', activeUserId);
  else localStorage.removeItem('activeUserId');
}, [activeUserId]);
  const [showMenu, setShowMenu] = useState(false);
  const [query, setQuery] = useState('');
  const [globalClientSearch, setGlobalClientSearch] = useState('');
  const [globalFocusedClientPhone, setGlobalFocusedClientPhone] = useState('');
  const [showGlobalSearchResults, setShowGlobalSearchResults] = useState(false);
  const [users, setUsers] = useState<User[]>(() => loadEmployeesFromStorage());
  const [products, setProducts] = useState<Product[]>(() => loadProductsFromStorage());
  const [orders, setOrders] = useState<ServiceOrder[]>(() => loadOrdersFromStorage());
  const [orderVersions, setOrderVersions] = useState<OrderVersion[]>(() => loadOrdersFromStorage().map((order) => ({
    id: uid('VER'),
    orderId: order.id,
    versionNo: order.currentVersion ?? 1,
    createdAt: order.intakeDate,
    createdBy: order.manager,
    reason: 'Початкова версія замовлення',
    snapshotData: JSON.stringify(order),
  })));
  const [customerList, setCustomerList] = useState<ClientRecord[]>(() => loadClientsFromStorage());
  const [contracts, setContracts] = useState<ContractRecord[]>(() => loadContractsFromStorage());
  const [contractActs, setContractActs] = useState<ContractActRecord[]>(() => loadContractActsFromStorage());
  const [bankImportItems, setBankImportItems] = useState<BankImportItem[]>(() => loadBankImportItemsFromStorage());
  const [bankImportDraft, setBankImportDraft] = useState<BankImportDraft | null>(null);
  const [sales, setSales] = useState<Sale[]>(initialSales);
  const [documents, setDocuments] = useState<PrintDocument[]>(initialDocuments);
  const [taxInvoices, setTaxInvoices] = useState<TaxInvoice[]>(() => loadTaxInvoicesFromStorage());
  const [simplePayments, setSimplePayments] = useState<SimpleOrderPaymentRecord[]>(() => loadSimplePaymentsFromStorage());
  const [clientNotifications, setClientNotifications] = useState<ClientNotification[]>(initialClientNotifications);
  const [internalMessages, setInternalMessages] = useState<InternalMessage[]>(initialInternalMessages);
  const [notificationTemplates] = useState<ClientNotificationTemplate[]>(clientNotificationTemplates);
  const [repairApprovals, setRepairApprovals] = useState<RepairApproval[]>(initialRepairApprovals);
  const [requirements, setRequirements] = useState<PartRequirement[]>(() => loadRequirementsFromStorage());
  const [purchases, setPurchases] = useState<PurchaseOrder[]>(() => loadPurchasesFromStorage());
  const [receipts, setReceipts] = useState<GoodsReceipt[]>(() => loadReceiptsFromStorage());
  const [movements, setMovements] = useState<StockMovement[]>(() => loadMovementsFromStorage());
  const [cashShift, setCashShift] = useState<CashShift>(() => loadCashShiftFromStorage());
  const [warehouseLocations] = useState<WarehouseLocation[]>(initialWarehouseLocations);
  const [orderMovementLogs, setOrderMovementLogs] = useState<OrderMovementLog[]>([]);
  const [orderUnits, setOrderUnits] = useState<OrderUnit[]>(initialOrderUnits);
  const [actionLogs, setActionLogs] = useState<ActionLog[]>(() => loadActionLogsFromStorage());
  const [backups, setBackups] = useState<BackupSnapshot[]>(() => loadBackupsFromStorage());
  const [selectedOrderId, setSelectedOrderId] = useState(() => loadOrdersFromStorage()[0]?.id ?? initialOrders[0]?.id ?? '');
  const [selectedSaleId, setSelectedSaleId] = useState(initialSales[0].id);
  const [selectedProductId, setSelectedProductId] = useState(() => loadProductsFromStorage()[0]?.id ?? initialProducts[0]?.id ?? '');
  const [selectedSaleProductId, setSelectedSaleProductId] = useState(() => loadProductsFromStorage()[0]?.id ?? initialProducts[0]?.id ?? '');
  const [qty, setQty] = useState(1);
  const [saleQty, setSaleQty] = useState(1);
  const [quickPhone, setQuickPhone] = useState('');
  const [quickClientDebtWarning, setQuickClientDebtWarning] = useState('');
  const [quickClientName, setQuickClientName] = useState('');
  const [quickDevice, setQuickDevice] = useState('');
  const [quickSerial, setQuickSerial] = useState('');
  const [quickProblem, setQuickProblem] = useState('');
  const [quickAppearance, setQuickAppearance] = useState('');
  const [quickEstimatedAmount, setQuickEstimatedAmount] = useState('');
  const [quickEngineerId, setQuickEngineerId] = useState('');
  const [quickContractId, setQuickContractId] = useState('');
  const [quickLocationCode, setQuickLocationCode] = useState('');
  const [quickComment, setQuickComment] = useState('');
  const [notice, setNotice] = useState('');
  const [toasts, setToasts] = useState<AppToast[]>([]);
  const [suggestedDocumentAction, setSuggestedDocumentAction] = useState<SuggestedDocumentAction | null>(null);
  const [dashboardFocus, setDashboardFocus] = useState<{ orderId: string; target: DashboardFocusTarget } | null>(null);
  const [loginPhone, setLoginPhone] = useState('');
  const [loginCode, setLoginCode] = useState('');
  const [loginError, setLoginError] = useState('');
  const [loginHint, setLoginHint] = useState('');
  const [loginChallenge, setLoginChallenge] = useState<LoginChallenge | null>(null);
  const [adminPreviewRole, setAdminPreviewRole] = useState<Role>('Адміністратор');
  const [adminPreviewUserId, setAdminPreviewUserId] = useState('');
  const [isNotificationsOpen, setIsNotificationsOpen] = useState(false);

  const sessionUserRecord = users.find((user) => user.id === sessionUserId) ?? null;
  const activeUserRecord = users.find((user) => user.id === activeUserId) ?? sessionUserRecord;
  const hookSafeUser = activeUserRecord ?? sessionUserRecord ?? SYSTEM_USER;
  const filteredOrders = (() => {
    const needle = query.trim();
    if (!needle) return orders;
    return orders.filter((order) => matchesOrderSearch(order, needle));
  })();
  const globalClientMatches = query.trim()
    ? customerList.filter((client) => matchesClientSearch(client, query))
    : [];
  const globalSearchResults = query.trim()
    ? [
        ...filteredOrders.slice(0, 5).map((order) => ({
          type: 'order' as const,
          id: order.id,
          title: `${order.id} · ${order.client}`,
          subtitle: `${order.device} · ${order.phone}`,
        })),
        ...globalClientMatches.slice(0, 5).map((client) => ({
          type: 'client' as const,
          id: client.phone,
          title: client.name,
          subtitle: `${client.phone} · ${client.taxId || 'клієнт'}`,
        })),
      ]
    : [];
  const analytics = (() => {
    const stockValue = products.reduce((sum, product) => sum + product.stock * product.cost, 0);
    const partsInRepairs = products.reduce((sum, product) => sum + product.reserved + product.withEngineer + product.installed, 0);
    const ordered = purchases.reduce((sum, purchase) => sum + purchase.items.reduce((inner, item) => inner + item.qty - item.received, 0), 0);
    const required = requirements.filter((item) => item.status === 'Потрібно' || item.status === 'До закупівлі').reduce((sum, item) => sum + item.qty, 0);
    const salesRevenue = sales.reduce((sum, sale) => sum + saleTotals(sale).paid, 0);
    const salesDebt = sales.reduce((sum, sale) => sum + saleTotals(sale).debt, 0);
    const serviceDebt = orders.reduce((sum, order) => sum + orderTotals(order).debt, 0);
    const salesProfit = sales.reduce((sum, sale) => sum + saleTotals(sale).profit, 0);
    const serviceProfit = orders.reduce((sum, order) => sum + orderTotals(order).finalProfit, 0);
    return { stockValue, partsInRepairs, ordered, required, salesRevenue, salesDebt, serviceDebt, salesProfit, serviceProfit };
  })();

  function stayOnCurrentPage(message = 'Перехід між розділами тимчасово доступний тільки через ліве меню.') {
    setNotice(message);
  }

  function openGlobalSearchResult(result: { type: 'order' | 'client'; id: string }) {
    if (result.type === 'order') {
      setSelectedOrderId(result.id);
      setDashboardFocus(null);
    } else {
      setGlobalFocusedClientPhone(result.id);
      setGlobalClientSearch(query.trim());
    }
    stayOnCurrentPage();
    setQuery('');
    setShowGlobalSearchResults(false);
  }

  function submitGlobalSearch() {
    const needle = query.trim();
    if (!needle) return;
    const exactOrder = findExactOrderMatch(orders, needle);
    if (exactOrder) {
      openGlobalSearchResult({ type: 'order', id: exactOrder.id });
      return;
    }
    const exactClient = findExactClientMatch(customerList, needle);
    if (exactClient) {
      openGlobalSearchResult({ type: 'client', id: exactClient.phone });
      return;
    }
    if (globalSearchResults.length === 1) {
      openGlobalSearchResult(globalSearchResults[0]);
      return;
    }
    if (globalSearchResults.length === 0) {
      setNotice('Нічого не знайдено.');
      setShowGlobalSearchResults(true);
      return;
    }
    setShowGlobalSearchResults(true);
    setNotice(`Знайдено результатів: ${globalSearchResults.length}`);
  }

  useEffect(() => {
    if (!activeUserId && sessionUserId) setActiveUserId(sessionUserId);
  }, [activeUserId, sessionUserId]);

  useEffect(() => {
    try { localStorage.removeItem(ACTION_LOGS_STORAGE_KEY); } catch {}
  }, []);

  useEffect(() => {
    if (!query.trim()) {
      setShowGlobalSearchResults(false);
    }
  }, [query]);

  useEffect(() => {
    if (sessionUserId && !sessionUserRecord) {
      setSessionUserId('');
      setActiveUserId('');
      setLoginChallenge(null);
      setLoginCode('');
      setLoginHint('');
      setLoginError('Сесію не знайдено. Увійдіть повторно.');
      return;
    }
    if (sessionUserRecord && !activeUserRecord) {
      setActiveUserId(sessionUserRecord.id);
    }
  }, [users, sessionUserId, sessionUserRecord, activeUserRecord]);

  useEffect(() => {
    if (!activeUserRecord || !['Адміністратор', 'Руководитель'].includes(activeUserRecord.role)) {
      if (adminPreviewRole !== 'Адміністратор') setAdminPreviewRole('Адміністратор');
      if (adminPreviewUserId) setAdminPreviewUserId('');
    }
  }, [activeUserRecord, adminPreviewRole, adminPreviewUserId]);

  useEffect(() => {
    if (!quickEngineerId || !users.some((user) => user.id === quickEngineerId && user.role === 'Інженер')) {
      setQuickEngineerId(users.find((user) => user.role === 'Інженер')?.id ?? '');
    }
  }, [users, quickEngineerId]);

  useEffect(() => {
    setRequirements((current) => {
      const next = [...current];
      const createdSignals: string[] = [];
      products.forEach((product) => {
        const needsMinimumPurchase = available(product) < Math.max(product.min, 1);
        if (!needsMinimumPurchase) return;
        const exists = next.some((requirement) =>
          requirement.productId === product.id
          && requirement.reason === 'Мінімум'
          && ['Потрібно', 'До закупівлі', 'Замовлено', 'В дорозі'].includes(requirement.status),
        );
        if (exists) return;
        next.unshift({
          id: uid('REQ'),
          orderId: '',
          productId: product.id,
          qty: Math.max(product.min - available(product), 1),
          status: 'Потрібно',
          reason: 'Мінімум',
          priority: available(product) <= 0 ? 'Високий' : 'Середній',
          requester: SYSTEM_USER.name,
          comment: 'Автоматичний сигнал: залишок нижче мінімального.',
        });
        createdSignals.push(`${product.sku} · ${product.name} · потрібно ${Math.max(product.min - available(product), 1)}`);
      });
      if (createdSignals.length > 0) {
        logSystemAction('Автосигнал закупівлі', 'warehouse', createdSignals.join('; '));
      }
      return next.length === current.length ? current : next;
    });
  }, [products]);

  useEffect(() => {
    setPurchases((current) => {
      const next = [...current];
      const draftedRequirementIds: string[] = [];
      const draftedProducts: string[] = [];
      requirements
        .filter((requirement) => requirement.reason === 'Мінімум' && requirement.status === 'Потрібно')
        .forEach((requirement) => {
          const alreadyExists = next.some((purchase) =>
            ['Нова', 'В роботі', 'Замовлено', 'В дорозі'].includes(purchase.status)
            && purchase.items.some((item) => item.requirementId === requirement.id || item.productId === requirement.productId),
          );
          if (alreadyExists) return;
          const product = products.find((item) => item.id === requirement.productId);
          next.unshift({
            id: uid('PO'),
            supplier: '',
            status: 'Нова',
            orderedAt: today,
            expectedAt: '',
            items: [{
              productId: requirement.productId,
              qty: requirement.qty,
              received: 0,
              price: product?.cost ?? 0,
              orderId: requirement.orderId || undefined,
              requirementId: requirement.id,
            }],
            reason: requirement.reason,
            priority: requirement.priority ?? 'Середній',
            comment: requirement.comment || 'Авточернетка закупки зі складського мінімуму.',
            requestedBy: SYSTEM_USER.name,
          });
          draftedRequirementIds.push(requirement.id);
          draftedProducts.push(`${product?.sku ?? requirement.productId} · ${product?.name ?? requirement.productId}`);
        });
      if (draftedRequirementIds.length > 0) {
        setRequirements((currentRequirements) => currentRequirements.map((requirement) => (
          draftedRequirementIds.includes(requirement.id)
            ? { ...requirement, status: 'До закупівлі' }
            : requirement
        )));
        logSystemAction('Авточернетка закупівлі', 'purchases', draftedProducts.join('; '));
      }
      return draftedRequirementIds.length > 0 ? next : current;
    });
  }, [products, requirements]);

  useEffect(() => {
    saveClientsToStorage(customerList);
  }, [customerList]);

  useEffect(() => {
    saveProductsToStorage(products);
  }, [products]);

  useEffect(() => {
    saveOrdersToStorage(orders);
  }, [orders]);

  useEffect(() => {
    saveSimplePaymentsToStorage(simplePayments);
  }, [simplePayments]);

  useEffect(() => {
    saveContractsToStorage(contracts);
  }, [contracts]);

  useEffect(() => {
    saveContractActsToStorage(contractActs);
  }, [contractActs]);

  useEffect(() => {
    saveBankImportItemsToStorage(bankImportItems);
  }, [bankImportItems]);

  useEffect(() => {
    saveRequirementsToStorage(requirements);
  }, [requirements]);

  useEffect(() => {
    savePurchasesToStorage(purchases);
  }, [purchases]);

  useEffect(() => {
    saveReceiptsToStorage(receipts);
  }, [receipts]);

  useEffect(() => {
    saveMovementsToStorage(movements);
  }, [movements]);

  useEffect(() => {
    saveTaxInvoicesToStorage(taxInvoices);
  }, [taxInvoices]);

  useEffect(() => {
    saveCashShiftToStorage(cashShift);
  }, [cashShift]);

  useEffect(() => {
    saveBackupsToStorage(backups);
  }, [backups]);

  useEffect(() => {
    const autoBackupIntervalMs = 12 * 60 * 60 * 1000;
    const ensureAutoBackup = () => {
      const latestAuto = backups.find((item) => item.source === 'auto');
      if (!latestAuto || Date.now() - latestAuto.createdAtTs >= autoBackupIntervalMs) {
        createBackupSnapshot('auto', 'Автоматична копія CRM', SYSTEM_USER.name);
        logSystemAction('Автобекап CRM', 'backup', 'Створено автоматичну резервну копію.');
      }
    };
    ensureAutoBackup();
    const timer = window.setInterval(ensureAutoBackup, 60 * 60 * 1000);
    return () => window.clearInterval(timer);
  }, [backups, users, customerList, products, orders, simplePayments, contracts, contractActs, bankImportItems, requirements, purchases, receipts, movements, taxInvoices, actionLogs, cashShift]);

  useEffect(() => {
    if (!orders.some((order) => order.id === selectedOrderId)) {
      setSelectedOrderId(orders[0]?.id ?? '');
    }
    if (!products.some((product) => product.id === selectedProductId)) {
      setSelectedProductId(products[0]?.id ?? '');
    }
    if (!products.some((product) => product.id === selectedSaleProductId)) {
      setSelectedSaleProductId(products[0]?.id ?? '');
    }
  }, [orders, products, selectedOrderId, selectedProductId, selectedSaleProductId]);

  useEffect(() => {
    if (!quickLocationCode) {
      const suggested = suggestFreeLocation('REPAIR');
      if (suggested) setQuickLocationCode(suggested);
    }
  }, [orders, quickLocationCode]);

  useEffect(() => {
    if (!notice) return;
    const toastId = uid('TOAST');
    const tone = inferToastTone(notice);
    setToasts((current) => [...current, { id: toastId, message: notice, tone }].slice(-5));
    const timeout = window.setTimeout(() => {
      setToasts((current) => current.filter((toast) => toast.id !== toastId));
    }, 2800);
    return () => window.clearTimeout(timeout);
  }, [notice]);

  async function handleClientsImport(file: File) {
    const rows = await readSpreadsheetRows(file);
    const importedClients = importClientsFromRows(rows);
    if (importedClients.length === 0) throw new Error('У файлі не знайдено клієнтів. Перевірте колонки: імʼя, телефон, email.');
    saveClientsToStorage(importedClients);
    setCustomerList(importedClients);
    setNotice(`Імпортовано клієнтів: ${importedClients.length}`);
  }

  async function handleProductsImport(file: File) {
    const rows = await readSpreadsheetRows(file);
    const importedProducts = importProductsFromRows(rows);
    if (importedProducts.length === 0) throw new Error('У файлі не знайдено товарів. Перевірте колонки: назва, кількість, ціна, артикул.');
    const mergedProducts = mergeImportedProducts(products, importedProducts);
    saveProductsToStorage(mergedProducts);
    setProducts(mergedProducts);
    setNotice(`Імпортовано партій: ${importedProducts.length}. Дублі поєднані з існуючою номенклатурою.`);
  }

  function clientTaxId(name: string, phone?: string) {
    return findClientRecord(customerList, name, phone, clientDetails(name, phone).edrpou)?.taxId
      || normalizeTaxId(clientDetails(name, phone).edrpou);
  }

  function clientOutstandingDebtByPhone(phone: string) {
    const normalizedPhone = phone.replace(/\D/g, '');
    return orders
      .filter((order) => order.phone.replace(/\D/g, '') === normalizedPhone)
      .reduce((sum, order) => sum + clientOrderDebt(order), 0);
  }

  function nextOrderStatusAfterAutoPayment(order: ServiceOrder, nextRemaining: number): OrderStatus {
    if (nextRemaining > 0.01 || order.status === 'Видано') return order.status;
    if (['Готовий до видачі', 'Очікує оплати', 'Очікує клієнта', 'Готово', 'Оплачено'].includes(order.status)) return 'Готовий до видачі';
    return order.status;
  }

  function debtOrdersForClient(input: { client: string; phone?: string; taxId?: string }) {
    const normalizedTaxId = normalizeTaxId(input.taxId);
    const normalizedPhone = normalizeImportedPhone(input.phone).replace(/\D/g, '');
    return orders
      .filter((order) => clientOrderDebt(order) > 0)
      .filter((order) => {
        const orderTaxId = clientTaxId(order.client, order.phone);
        if (normalizedTaxId && orderTaxId && normalizeTaxId(orderTaxId) === normalizedTaxId) return true;
        if (normalizedPhone && order.phone.replace(/\D/g, '') === normalizedPhone) return true;
        return order.client.trim().toLowerCase() === input.client.trim().toLowerCase();
      })
      .sort((a, b) => {
        const aTime = parseDateTime(a.debtSince ?? a.issuedDebtAt ?? a.intakeDate)?.getTime() ?? 0;
        const bTime = parseDateTime(b.debtSince ?? b.issuedDebtAt ?? b.intakeDate)?.getTime() ?? 0;
        return aTime - bTime;
      });
  }

  function distributeClientDebtPayment(paymentRecord: SimpleOrderPaymentRecord) {
    const debtOrders = debtOrdersForClient({
      client: paymentRecord.client,
      taxId: paymentRecord.clientTaxId ?? paymentRecord.taxId,
    });
    if (debtOrders.length === 0) {
      return { updatedOrders: orders, allocations: [] as Array<{ orderId: string; amount: number; remainingAfter: number }>, unallocated: paymentRecord.amount };
    }
    let remainingAmount = paymentRecord.amount;
    const allocations: Array<{ orderId: string; amount: number; remainingAfter: number }> = [];
    const updatedOrders = orders.map((order) => {
      const targetDebtOrder = debtOrders.find((entry) => entry.id === order.id);
      if (!targetDebtOrder || remainingAmount <= 0) return order;
      const currentDebt = clientOrderDebt(order);
      if (currentDebt <= 0) return order;
      const appliedAmount = Math.min(currentDebt, remainingAmount);
      remainingAmount -= appliedAmount;
      const debtAfter = Math.max(currentDebt - appliedAmount, 0);
      const nextStatus = nextOrderStatusAfterAutoPayment(order, debtAfter);
      allocations.push({ orderId: order.id, amount: appliedAmount, remainingAfter: debtAfter });
      const allocatedPayment: Payment = {
        id: uid('PAY-DEBT'),
        date: paymentRecord.date || today,
        amount: appliedAmount,
        method: mapSimpleMethodToPaymentMethod(paymentRecord.method),
        type: appliedAmount >= currentDebt - 0.01 ? 'Повна оплата' : 'Часткова оплата',
        transactionNo: paymentRecord.documentRef || paymentRecord.id,
        acceptedBy: activeUser.name,
        status: 'Підтверджено',
        confirmedBy: activeUser.name,
        confirmedAt: today,
        orderId: order.id,
        comment: `Розподілено по боргу клієнта · ${paymentRecord.payer || paymentRecord.client}`,
      };
      return {
        ...order,
        status: nextStatus,
        statusChangedAt: nextStatus !== order.status ? today : order.statusChangedAt,
        payments: [allocatedPayment, ...order.payments],
        activityLog: [
          {
            id: uid('ACT'),
            date: today,
            action: 'Погашення боргу',
            detail: `${money(appliedAmount)} · платіж ${paymentRecord.date || today} · ${order.id} · залишок ${money(debtAfter)}`,
          },
          ...(order.activityLog ?? []),
        ],
      };
    });
    return { updatedOrders, allocations, unallocated: remainingAmount };
  }

  function recalcContractActs(nextPayments: SimpleOrderPaymentRecord[], actsSource = contractActs) {
    return actsSource.map((act) => {
      const paidAmount = nextPayments
        .filter((payment) => payment.status === 'подтвержден' && payment.actId === act.id)
        .reduce((sum, payment) => sum + payment.amount, 0);
      const remainingAmount = Math.max(act.amount - paidAmount, 0);
      return {
        ...act,
        paidAmount,
        remainingAmount,
        status: recalculateActStatus({ amount: act.amount, paidAmount }),
      };
    });
  }

  function clientBalanceAmount(client: ClientRecord) {
    const normalizedTaxId = normalizeTaxId(client.taxId);
    const incoming = simplePayments
      .filter((payment) => payment.status === 'подтвержден')
      .filter((payment) => {
        const sameTaxId = normalizedTaxId && normalizeTaxId(payment.clientTaxId ?? payment.taxId) === normalizedTaxId;
        const sameClient = payment.client.trim().toLowerCase() === client.name.trim().toLowerCase();
        return sameTaxId || sameClient;
      })
      .reduce((sum, payment) => sum + payment.amount, 0);
    const closedActsAmount = contractActs
      .filter((act) => act.status === 'Оплачен')
      .filter((act) => {
        const contract = contracts.find((item) => item.id === act.contractId);
        return contract?.client.trim().toLowerCase() === client.name.trim().toLowerCase();
      })
      .reduce((sum, act) => sum + act.amount, 0);
    return incoming - closedActsAmount;
  }

  function requestPhoneCode() {
    const normalized = normalizePhone(loginPhone);
    const employees = loadEmployeesFromStorage();
    const searchableEmployees = employees.length > 0 ? employees : users;
    console.log('LOGIN employees', searchableEmployees);
    console.log('LOGIN phone', normalized);
    const user = searchableEmployees.find((item) => normalizePhone(item.phone ?? '') === normalized);
    if (!normalized) {
      setLoginError('Вкажіть телефон співробітника.');
      setLoginHint('');
      return;
    }
    if (!user) {
      setLoginError('Телефон не найден в базе сотрудников');
      setLoginHint('');
      return;
    }
    if (user.session === 'Заблокована') {
      setLoginError('Цей співробітник заблокований. Вхід заборонено.');
      setLoginHint('');
      return;
    }
    const nowTs = Date.now();
    const code = '1111';
    setLoginChallenge({ userId: user.id, code, phone: user.phone ?? loginPhone, expiresAt: nowTs + loginCodeLifetimeMs, attempts: 0 });
    setLoginCode('');
    setLoginError('');
    setLoginHint(`Тестовий код для ${user.name}: ${code}`);
    prependActionLog({ id: uid('LOG'), date: today, user: user.name, role: user.role, action: 'Запит коду входу', entity: 'Авторизація', comment: 'Створено тестовий одноразовий код 1111 без SMS-інтеграції.' });
  }

  function verifyPhoneCode() {
    if (!loginChallenge) return;
    const user = users.find((item) => item.id === loginChallenge.userId);
    if (!user) {
      setLoginError('Співробітника більше немає в базі.');
      setLoginHint('');
      return;
    }
    if (user.session === 'Заблокована') {
      setLoginError('Цей співробітник заблокований. Вхід заборонено.');
      setLoginHint('');
      return;
    }
    if (Date.now() > loginChallenge.expiresAt) {
      setLoginChallenge(null);
      setLoginCode('');
      setLoginError('Термін дії коду минув. Запросіть новий код.');
      setLoginHint('');
      return;
    }
    if (loginCode.trim() !== loginChallenge.code) {
      const nextAttempts = loginChallenge.attempts + 1;
      if (nextAttempts >= 3) {
        setLoginChallenge(null);
        setLoginCode('');
        setLoginError('Перевищено кількість спроб. Запросіть новий код.');
        setLoginHint('');
        return;
      }
      setLoginChallenge((current) => (current ? { ...current, attempts: nextAttempts } : current));
      setLoginError(`Невірний код. Залишилось спроб: ${3 - nextAttempts}.`);
      return;
    }
    setSessionUserId(user.id);
    setActiveUserId(user.id);
    setPage(defaultPageForUser(user));
    setShowMenu(false);
    setLoginChallenge(null);
    setLoginCode('');
    setLoginPhone('');
    setLoginError('');
    setLoginHint('');
    prependActionLog({ id: uid('LOG'), date: today, user: user.name, role: user.role, action: 'Вхід у систему', entity: 'Сесія', comment: 'Вхід по телефону і одноразовому коду.' });
  }

  function resetPhoneLogin() {
    setLoginChallenge(null);
    setLoginCode('');
    setLoginError('');
    setLoginHint('');
  }

  function logoutEmployee() {
    if (activeUser) {
      prependActionLog({ id: uid('LOG'), date: today, user: activeUser.name, role: activeUser.role, action: 'Вихід із системи', entity: 'Сесія', comment: 'Сесію завершено користувачем.' });
    }
    setAdminPreviewRole('Адміністратор');
    setAdminPreviewUserId('');
    setSessionUserId('');
    setActiveUserId('');
    setLoginChallenge(null);
    setLoginCode('');
    setLoginPhone('');
    setLoginError('');
    setLoginHint('');
  }

  const sessionUser = sessionUserRecord;
  const activeUser = activeUserRecord ?? sessionUserRecord ?? hookSafeUser;
  const isAdminPreviewAvailable = ['Адміністратор', 'Руководитель'].includes(activeUser.role);
  const currentRole = adminPreviewRole;
  const previewUsers = users.filter((user) => user.role === currentRole);
  const selectedPreviewUser = previewUsers.find((user) => user.id === adminPreviewUserId) ?? null;
  const currentUser = selectedPreviewUser ?? (currentRole === activeUser.role ? activeUser : null);
  const viewUser = isAdminPreviewAvailable
    ? (currentUser ? normalizeUserRecord(currentUser) : buildPreviewUser(currentRole, activeUser))
    : activeUser;

  const canSwitchUsers = sessionUser?.role === 'Руководитель' && sessionUser.session === 'Активна';
  const isOwnerControlView = sessionUser?.role === 'Руководитель' && activeUser.id !== sessionUser.id;
  const canViewPage = (targetPage: Page) => canRoleViewPage(viewUser, targetPage);
  const canDo = (permission: Permission) => roleFinePermissions[viewUser.role].includes(permission) || hasPermissionOverride(viewUser, permission);
  const visibleNavItems = navItems
    .filter((item) => canViewPage(item.id))
    .filter((item) => !(viewUser.role === 'Менеджер' && item.id === 'orders'));
  const hideGlobalSearch = viewUser.role === 'Менеджер' && (page === 'orders' || page === 'my-orders');
  const repairOrders = filteredOrders;
  const engineerOwnOrders = viewUser.role === 'Інженер' ? orders.filter((order) => getEngineerAssignedUserId(order, users) === viewUser.id) : orders;
  const managerOwnOrders = viewUser.role === 'Менеджер' ? orders.filter((order) => getManagerUserId(order, users) === viewUser.id || order.manager === viewUser.name) : orders;
  const engineerHiddenAfterReady: OrderStatus[] = ['Готовий до видачі', 'Не підлягає ремонту', 'Видано', 'Закрито', 'Скасовано'];
  const myOrders = viewUser.role === 'Інженер'
    ? engineerOwnOrders.filter((order) => !engineerHiddenAfterReady.includes(order.status))
    : viewUser.role === 'Менеджер'
      ? managerOwnOrders
      : orders;
  const selectedRepairOrder = repairOrders.find((order) => order.id === selectedOrderId) ?? repairOrders[0] ?? orders[0];
  const selectedMyOrder = myOrders.find((order) => order.id === selectedOrderId) ?? myOrders[0] ?? selectedRepairOrder;
  const selectedOrder = page === 'my-orders' ? selectedMyOrder : selectedRepairOrder;
  const selectedSale = sales.find((sale) => sale.id === selectedSaleId) ?? sales[0];
  const visibleInternalMessages = viewUser.role === 'Руководитель'
    ? internalMessages
    : internalMessages.filter((message) => message.toUser === viewUser.name || message.toRole === viewUser.role);
  const unreadInternalMessages = visibleInternalMessages.filter((message) => message.status === 'Нове').length;

  function productName(productId: string) {
    return products.find((product) => product.id === productId)?.name ?? 'Невідома запчастина';
  }

  function getLocationByCode(code?: string) {
    return warehouseLocations.find((location) => location.code === code);
  }

  function locationZone(code?: string) {
    return getLocationByCode(code)?.zone ?? (code === 'OUT' ? 'OUT' : null);
  }

  function suggestFreeLocation(zone: WarehouseZone) {
    const busyLocations = new Set(
      orders
        .filter((order) => order.locationCode)
        .map((order) => order.locationCode as string),
    );
    return warehouseLocations.find((location) => location.zone === zone && !busyLocations.has(location.code))?.code;
  }

  function moveOrder(orderId: string, newLocation: string) {
    const targetLocation = getLocationByCode(newLocation);
    if (!targetLocation && newLocation !== 'OUT') {
      setNotice('Нова локація не знайдена.');
      return;
    }
    const order = orders.find((item) => item.id === orderId);
    if (!order) return;
    if (order.locationCode === newLocation) {
      setNotice(`Замовлення вже знаходиться у ${newLocation}.`);
      return;
    }
    if (newLocation !== 'OUT' && orders.some((item) => item.id !== orderId && item.locationCode === newLocation)) {
      setNotice(`Комірка ${newLocation} вже зайнята іншим замовленням.`);
      return;
    }
    setOrders((current) => current.map((item) => (
      item.id !== orderId
        ? item
        : {
            ...item,
            locationCode: newLocation === 'OUT' ? undefined : newLocation,
            locationStatus: newLocation === 'OUT' ? 'Видано' : item.locationStatus === 'У інженера' ? 'У інженера' : 'У комірці',
            statusHistory: [
              {
                id: uid('H'),
                oldStatus: item.status,
                newStatus: item.status,
                changedBy: activeUser.name,
                changedAt: today,
                comment: `Переміщення по складу: ${item.locationCode ?? 'без місця'} -> ${newLocation}.`,
              },
              ...item.statusHistory,
            ],
          }
    )));
    setOrderUnits((current) => current.map((unit) => (
      unit.orderId !== orderId
        ? unit
        : {
            ...unit,
            locationCode: newLocation === 'OUT' ? undefined : newLocation,
            status: targetLocation?.zone === 'READY' ? 'На полці' : unit.status,
            lastActionAt: today,
          }
    )));
    setOrderMovementLogs((current) => [
      {
        id: uid('MOVE'),
        orderId,
        fromLocation: order.locationCode,
        toLocation: newLocation,
        userId: activeUser.id,
        timestamp: today,
      },
      ...current,
    ]);
    logAction('Переміщення замовлення', orderId, `${order.locationCode ?? 'без місця'} -> ${newLocation}.`);
    setNotice(`${orderId}: переміщено ${order.locationCode ?? 'без місця'} -> ${newLocation}.`);
  }

  function addMovement(movement: Omit<StockMovement, 'id' | 'date'>) {
    setMovements((current) => [{ id: uid('MOV'), date: today, actor: movement.actor ?? activeUser.name, ...movement }, ...current]);
  }

  function prependActionLog(_entry: ActionLog) {
    // Temporarily disabled to prevent crm_action_logs growth and storage quota errors.
  }

  function logAction(action: string, entity: string, comment: string) {
    prependActionLog({ id: uid('LOG'), date: today, user: activeUser.name, role: activeUser.role, action, entity, comment });
  }

  function logSystemAction(action: string, entity: string, comment: string) {
    prependActionLog({ id: uid('LOG'), date: today, user: SYSTEM_USER.name, role: SYSTEM_USER.role, action, entity, comment });
  }

  function logRiskConfirmation(order: ServiceOrder, action: string, risks: string[]) {
    risks.forEach((risk) => {
      logAction('Підтвердження ризику', order.id, `${action}. Ризик: ${risk}`);
    });
  }

  function ensureOrderEditable(order: ServiceOrder, actionLabel: string) {
    if (!isOrderDealClosed(order, documents, taxInvoices)) return true;
    setNotice(`Угоду ${order.id} вже закрито. Дія "${actionLabel}" недоступна.`);
    return false;
  }

  function buildBackupPayload(): BackupPayload {
    return {
      version: 1,
      users,
      clients: customerList,
      products,
      orders,
      payments: simplePayments,
      contracts,
      contractActs,
      bankImportItems,
      requirements,
      purchases,
      receipts,
      movements,
      taxInvoices,
      actionLogs,
      cashShift,
    };
  }

  function normalizeBackupRetention(items: BackupSnapshot[]) {
    const cutoff = Date.now() - (30 * 24 * 60 * 60 * 1000);
    return items
      .filter((item) => item.createdAtTs >= cutoff)
      .sort((a, b) => b.createdAtTs - a.createdAtTs)
      .slice(0, 60);
  }

  function createBackupSnapshot(source: BackupSnapshot['source'], label: string, createdBy: string) {
    const snapshot: BackupSnapshot = {
      id: uid('BKP'),
      createdAt: today,
      createdAtTs: Date.now(),
      createdBy,
      source,
      label,
      payload: buildBackupPayload(),
    };
    setBackups((current) => normalizeBackupRetention([snapshot, ...current]));
    return snapshot;
  }

  function restoreBackupSnapshot(snapshotId: string) {
    const snapshot = backups.find((item) => item.id === snapshotId);
    if (!snapshot) {
      setNotice('Резервну копію не знайдено.');
      return;
    }
    if (typeof window !== 'undefined' && !window.confirm(`Відновити CRM із копії ${snapshot.createdAt}?`)) return;
    const payload = snapshot.payload;
    saveEmployeesToStorage(payload.users);
    saveClientsToStorage(payload.clients);
    saveProductsToStorage(payload.products);
    saveOrdersToStorage(payload.orders);
    saveSimplePaymentsToStorage(payload.payments);
    saveContractsToStorage(payload.contracts);
    saveContractActsToStorage(payload.contractActs);
    saveBankImportItemsToStorage(payload.bankImportItems);
    saveRequirementsToStorage(payload.requirements);
    savePurchasesToStorage(payload.purchases);
    saveReceiptsToStorage(payload.receipts);
    saveMovementsToStorage(payload.movements);
    saveTaxInvoicesToStorage(payload.taxInvoices ?? []);
    saveCashShiftToStorage(payload.cashShift);
    const restoredAudit = [] as ActionLog[];
    setUsers(payload.users);
    setCustomerList(payload.clients);
    setProducts(payload.products);
    setOrders(payload.orders);
    setSimplePayments(payload.payments);
    setContracts(payload.contracts);
    setContractActs(payload.contractActs);
    setBankImportItems(payload.bankImportItems);
    setRequirements(payload.requirements);
    setPurchases(payload.purchases);
    setReceipts(payload.receipts);
    setMovements(payload.movements);
    setTaxInvoices((payload.taxInvoices ?? []).map(normalizeStoredTaxInvoice));
    setCashShift(payload.cashShift ?? initialCashShift);
    setSelectedOrderId(payload.orders[0]?.id ?? '');
    setSelectedProductId(payload.products[0]?.id ?? '');
    setSelectedSaleProductId(payload.products[0]?.id ?? '');
    setNotice(`CRM відновлено з копії ${snapshot.createdAt}.`);
  }

  function exportBackupSnapshot(snapshotId: string) {
    const snapshot = backups.find((item) => item.id === snapshotId);
    if (!snapshot) {
      setNotice('Резервну копію не знайдено.');
      return;
    }
    downloadExportFile(`crm-backup-${snapshot.createdAtTs}.json`, 'application/json', JSON.stringify(snapshot, null, 2));
    logAction('Експорт backup', snapshot.id, `Експортовано копію від ${snapshot.createdAt}.`);
    setNotice('Резервну копію експортовано у JSON.');
  }

  function exportCurrentLiveData() {
    const payload = buildBackupPayload();
    const exportEnvelope = {
      type: 'crm-live-data-export',
      version: 1,
      exportedAt: today,
      exportedAtTs: Date.now(),
      exportedBy: activeUser.name,
      payload,
    };
    downloadExportFile(`crm-live-data-${exportEnvelope.exportedAtTs}.json`, 'application/json', JSON.stringify(exportEnvelope, null, 2));
    logAction('Експорт живих даних', 'CRM', 'Експортовано актуальні дані CRM з браузерного сховища.');
    setNotice('Живі дані CRM експортовано у JSON.');
  }

  async function importBackupFile(file: File) {
    try {
      const text = await file.text();
      const parsed = JSON.parse(text) as BackupSnapshot | BackupPayload;
      const snapshot: BackupSnapshot = 'payload' in parsed
        ? parsed
        : {
            id: uid('BKP'),
            createdAt: today,
            createdAtTs: Date.now(),
            createdBy: activeUser.name,
            source: 'import',
            label: 'Імпортована копія',
            payload: parsed,
          };
      if (!snapshot.payload?.orders || !snapshot.payload?.products || !snapshot.payload?.clients) {
        setNotice('Файл backup не розпізнано.');
        return;
      }
      setBackups((current) => normalizeBackupRetention([{ ...snapshot, id: uid('BKP'), source: 'import', createdAt: snapshot.createdAt ?? today, createdAtTs: snapshot.createdAtTs ?? Date.now() }, ...current]));
      logAction('Імпорт backup', file.name, 'Резервну копію додано в список відновлення.');
      setNotice('Резервну копію імпортовано. Тепер її можна відновити зі списку.');
    } catch {
      setNotice('Не вдалося імпортувати резервну копію.');
    }
  }

  function downloadExportFile(filename: string, mimeType: string, content: string) {
    const blob = new Blob([content], { type: `${mimeType};charset=utf-8` });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }

  function exportRowsAsCsv(rows: Array<Record<string, string | number>>) {
    if (rows.length === 0) return '';
    const headers = Object.keys(rows[0]);
    const escapeCell = (value: string | number) => `"${String(value ?? '').replace(/"/g, '""')}"`;
    return [headers.join(';'), ...rows.map((row) => headers.map((header) => escapeCell(row[header] ?? '')).join(';'))].join('\n');
  }

  function exportRowsAsExcelHtml(title: string, rows: Array<Record<string, string | number>>) {
    if (rows.length === 0) return `<html><head><meta charset="utf-8" /></head><body><h2>${escapeHtml(title)}</h2><p>Немає даних за обраний період.</p></body></html>`;
    const headers = Object.keys(rows[0]);
    const thead = `<tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join('')}</tr>`;
    const tbody = rows.map((row) => `<tr>${headers.map((header) => `<td>${escapeHtml(String(row[header] ?? ''))}</td>`).join('')}</tr>`).join('');
    return `<html><head><meta charset="utf-8" /><style>body{font-family:Arial,sans-serif;font-size:12px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #999;padding:6px 8px;text-align:left}th{background:#f1f5f9}</style></head><body><h2>${escapeHtml(title)}</h2><table><thead>${thead}</thead><tbody>${tbody}</tbody></table></body></html>`;
  }

  function exportAccounting(kind: AccountingExportKind, format: AccountingExportFormat, period: AccountingPeriod) {
    const paymentMethodLabel = (methods: PaymentMethod[]) => methods.length ? Array.from(new Set(methods)).join(', ') : 'Без оплати';
    const exportRows: Array<Record<string, string | number>> = kind === 'sales'
      ? [
        ...orders
          .filter((order) => isDateInPeriod(order.payments[order.payments.length - 1]?.date ?? order.intakeDate, period))
          .map((order) => {
            const totals = orderTotals(order);
      const paidPayments = order.payments.filter(paymentCountsAsApplied);
            return {
              'Дата': extractDayKey(paidPayments[paidPayments.length - 1]?.date ?? order.intakeDate),
              'Номер заказа': order.id,
              'Клієнт': order.client,
              'Роботи': totals.works,
              'Запчастини': totals.parts,
              'Разом': totals.total,
              'Оплачено': totals.paid,
              'Оплата': paymentMethodLabel(paidPayments.map((payment) => payment.method)),
              'Статус оплати': simpleRepairPaymentStatus(order),
            };
          }),
        ...sales
          .filter((sale) => isDateInPeriod(sale.payments[sale.payments.length - 1]?.date ?? sale.date, period))
          .map((sale) => {
            const totals = saleTotals(sale);
            return {
              'Дата': extractDayKey(sale.payments[sale.payments.length - 1]?.date ?? sale.date),
              'Номер заказа': sale.id,
              'Клієнт': sale.client,
              'Роботи': 0,
              'Запчастини': totals.total,
              'Разом': totals.total,
              'Оплачено': totals.paid,
              'Оплата': paymentMethodLabel(sale.payments.filter(paymentCountsAsApplied).map((payment) => payment.method)),
              'Статус оплати': totals.debt > 0 ? (totals.paid > 0 ? 'Частково оплачено' : 'Не оплачено') : 'Оплачено',
            };
          }),
      ]
      : kind === 'purchases'
        ? receipts
          .filter((receipt) => isDateInPeriod(receipt.date, period))
          .map((receipt) => ({
            'Дата': extractDayKey(receipt.date),
            'Товар': productName(receipt.productId),
            'Кількість': receipt.qty,
            'Ціна закупки': receipt.price,
            'Постачальник': receipt.supplier,
          }))
        : kind === 'stock'
          ? movements
            .filter((movement) => isDateInPeriod(movement.date, period))
            .map((movement) => ({
              'Дата': extractDayKey(movement.date),
              'Тип': movement.type,
              'Товар': productName(movement.productId),
              'Кількість': movement.qty,
              'Документ': movement.orderId ?? movement.purchaseId ?? '',
              'Підстава': movement.basis ?? '',
              'Хто': movement.actor ?? '',
              'Залишок зараз': products.find((product) => product.id === movement.productId)?.stock ?? 0,
              'Коментар': movement.comment,
            }))
          : (() => {
            const periodOrderPayments = orders.flatMap((order) => order.payments.filter((payment) => paymentCountsAsApplied(payment) && isDateInPeriod(payment.date, period)));
            const periodSalePayments = sales.flatMap((sale) => sale.payments.filter((payment) => paymentCountsAsApplied(payment) && isDateInPeriod(payment.date, period)));
            const revenue = [...periodOrderPayments, ...periodSalePayments].reduce((sum, payment) => sum + payment.amount, 0);
            const purchasesAmount = receipts.filter((receipt) => isDateInPeriod(receipt.date, period)).reduce((sum, receipt) => sum + (receipt.qty * receipt.price), 0);
            const debts = orders.reduce((sum, order) => sum + orderTotals(order).debt, 0) + sales.reduce((sum, sale) => sum + saleTotals(sale).debt, 0);
            const unpaidOrdersCount = orders.filter((order) => orderTotals(order).debt > 0).length;
            return [
              { 'Період': period === 'today' ? 'Сьогодні' : period === 'week' ? 'Тиждень' : 'Місяць', 'Показник': 'Надходження грошей', 'Значення': revenue },
              { 'Період': period === 'today' ? 'Сьогодні' : period === 'week' ? 'Тиждень' : 'Місяць', 'Показник': 'Закупки', 'Значення': purchasesAmount },
              { 'Період': period === 'today' ? 'Сьогодні' : period === 'week' ? 'Тиждень' : 'Місяць', 'Показник': 'Борги клієнтів', 'Значення': debts },
              { 'Період': period === 'today' ? 'Сьогодні' : period === 'week' ? 'Тиждень' : 'Місяць', 'Показник': 'Неоплачені замовлення', 'Значення': unpaidOrdersCount },
            ];
          })();

    const titleMap: Record<AccountingExportKind, string> = {
      sales: 'Продажі для бухгалтерії',
      purchases: 'Закупки для бухгалтерії',
      stock: 'Рух складу для бухгалтерії',
      summary: 'Зведення за період',
    };
    const periodSuffix = period === 'today' ? 'today' : period === 'week' ? 'week' : 'month';
    const filenameBase = `accounting-${kind}-${periodSuffix}`;

    if (format === 'json') {
      downloadExportFile(`${filenameBase}.json`, 'application/json', JSON.stringify(exportRows, null, 2));
    } else if (format === 'csv') {
      downloadExportFile(`${filenameBase}.csv`, 'text/csv', exportRowsAsCsv(exportRows));
    } else {
      downloadExportFile(`${filenameBase}.xls`, 'application/vnd.ms-excel', exportRowsAsExcelHtml(titleMap[kind], exportRows));
    }

    logAction('Вигрузка в бухгалтерію', filenameBase, `${titleMap[kind]} · формат ${format} · період ${period}.`);
    setNotice(`${titleMap[kind]} (${format.toUpperCase()}) підготовлено.`);
  }

  function createOrderVersion(order: ServiceOrder, reason: string) {
    const versionNo = Math.max(0, ...orderVersions.filter((item) => item.orderId === order.id).map((item) => item.versionNo)) + 1;
    const nextSnapshot = { ...order, currentVersion: versionNo };
    setOrderVersions((current) => [
      {
        id: uid('VER'),
        orderId: order.id,
        versionNo,
        createdAt: today,
        createdBy: activeUser.name,
        reason,
        snapshotData: JSON.stringify(nextSnapshot),
      },
      ...current,
    ]);
    setOrders((current) => current.map((item) => (item.id === order.id ? { ...item, currentVersion: versionNo } : item)));
    logAction('Нова версія замовлення', order.id, `Версія ${versionNo}: ${reason}.`);
    return versionNo;
  }

  function persistOrdersUpdate(updater: (current: ServiceOrder[]) => ServiceOrder[]) {
    setOrders((current) => {
      const updatedOrders = updater(current);
      saveOrdersToStorage(updatedOrders);
      return updatedOrders;
    });
  }

  function patchSimpleManagerOrder(orderId: string, patch: Partial<ServiceOrder>) {
    const targetOrder = orders.find((order) => order.id === orderId);
    if (!targetOrder) return;
    if (!ensureOrderEditable(targetOrder, 'Редагування замовлення')) return;
    const activityEntries: Array<{ action: string; detail: string }> = [];
    const actionEntries: Array<{ action: string; detail: string }> = [];
    if (Object.prototype.hasOwnProperty.call(patch, 'engineer') && patch.engineer !== targetOrder.engineer) {
      activityEntries.push({ action: 'Інженера призначено', detail: patch.engineer ? `${patch.engineer}` : 'Інженера знято' });
      actionEntries.push({ action: 'Призначення інженера', detail: `${targetOrder.engineer || 'Не призначено'} -> ${patch.engineer || 'Не призначено'}` });
    }
    if (Object.prototype.hasOwnProperty.call(patch, 'repairPrice') && patch.repairPrice !== targetOrder.repairPrice) {
      activityEntries.push({ action: 'Суму оновлено', detail: `${money(targetOrder.repairPrice ?? 0)} -> ${money(patch.repairPrice ?? 0)}` });
      actionEntries.push({ action: 'Зміна суми ремонту', detail: `${money(targetOrder.repairPrice ?? 0)} -> ${money(patch.repairPrice ?? 0)}` });
    }
    if (Object.prototype.hasOwnProperty.call(patch, 'intakeComment') && (patch.intakeComment ?? '') !== (targetOrder.intakeComment ?? '')) {
      activityEntries.push({ action: 'Коментар оновлено', detail: patch.intakeComment?.trim() ? patch.intakeComment.trim() : 'Коментар очищено' });
      actionEntries.push({ action: 'Коментар менеджера', detail: patch.intakeComment?.trim() ? 'Коментар оновлено.' : 'Коментар очищено.' });
    }
    persistOrdersUpdate((current) => current.map((item) => (
      item.id === orderId
        ? {
            ...item,
            ...patch,
            activityLog: activityEntries.length > 0
              ? [
                  ...activityEntries.map((entry) => ({ id: uid('ACT'), date: today, action: entry.action, detail: entry.detail })),
                  ...(item.activityLog ?? []),
                ]
              : item.activityLog,
          }
        : item
    )));
    actionEntries.forEach((entry) => logAction(entry.action, orderId, entry.detail));
  }

  function appendSimpleOrderActivity(orderId: string, action: string, detail?: string) {
    persistOrdersUpdate((current) => current.map((item) => (
      item.id === orderId
        ? {
            ...item,
            activityLog: [
              { id: uid('ACT'), date: today, action, detail },
              ...(item.activityLog ?? []),
            ],
          }
        : item
    )));
  }

  function updateSimpleManagerOrderStatus(orderId: string, nextStatus: 'В ремонті' | 'Готово' | 'Видано') {
    const targetOrder = orders.find((order) => order.id === orderId);
    if (!targetOrder) return;
    if (!ensureOrderEditable(targetOrder, `Зміна статусу на "${nextStatus}"`)) return;
    const total = Math.max(targetOrder.repairPrice ?? targetOrder.estimatedAmount ?? orderTotals(targetOrder).total, 0);
    const paid = targetOrder.payments.filter(paymentCountsAsApplied).reduce((sum, payment) => sum + payment.amount, 0);
    const remaining = Math.max(total - paid, 0);
    if (nextStatus === 'В ремонті') {
      if (!targetOrder.engineer) {
        setNotice('Спочатку призначте інженера.');
        return;
      }
      if (targetOrder.status !== 'Прийнято') {
        setNotice('У роботу можна перевести тільки нове замовлення після призначення інженера.');
        return;
      }
    }
    if (nextStatus === 'Готово') {
      if (simpleRepairStatus(targetOrder.status) !== 'В ремонті') {
        setNotice('Позначити як готове можна тільки замовлення у роботі.');
        return;
      }
      if (!targetOrder.engineerWorkCompletedAt && !targetOrder.returnedToCellAt) {
        setNotice('Інженер ще не завершив замовлення.');
        return;
      }
    }
    if (nextStatus === 'Видано') {
      if (targetOrder.status !== 'Готовий до видачі') {
        setNotice('Видати можна тільки зі статусу "Готовий до видачі".');
        return;
      }
      if (remaining > 0) {
        setNotice('Неможливо видати — є заборгованість');
        return;
      }
    }
    const statusToStore: OrderStatus = nextStatus === 'Готово'
      ? (remaining > 0 ? 'Очікує оплати' : 'Готовий до видачі')
      : nextStatus;
    const issuedDebtAmount = undefined;
    persistOrdersUpdate((current) => current.map((item) => (
      item.id === orderId
        ? {
            ...item,
            status: statusToStore,
            statusChangedAt: today,
            issuedAt: nextStatus === 'Видано' ? today : item.issuedAt,
            issuedDebtAmount: issuedDebtAmount ?? item.issuedDebtAmount,
            issuedDebtAt: issuedDebtAmount ? today : item.issuedDebtAt,
            issuedDebtManager: issuedDebtAmount ? activeUser.name : item.issuedDebtManager,
            debtSince: issuedDebtAmount ? today : item.debtSince,
            statusHistory: [
              { id: uid('H'), oldStatus: item.status, newStatus: statusToStore, changedBy: activeUser.name, changedAt: today, comment: `Менеджер змінив статус на "${statusToStore}".` },
              ...(item.statusHistory ?? []),
            ],
          }
        : item
    )));
    appendSimpleOrderActivity(orderId, 'Статус', statusToStore);
    logAction(
      nextStatus === 'Видано' ? 'Видача замовлення' : nextStatus === 'Готово' ? 'Завершення ремонту' : 'Зміна статусу замовлення',
      orderId,
      `${targetOrder.status} -> ${statusToStore}.`,
    );
    if (nextStatus === 'В ремонті') {
      enqueueClientNotification({ ...targetOrder, status: statusToStore }, 'Ремонт розпочато');
    }
    if (nextStatus === 'Готово') {
      enqueueClientNotification(
        { ...targetOrder, status: statusToStore },
        statusToStore === 'Готовий до видачі' ? 'Готово до видачі' : 'Очікує оплату',
      );
    }
    if (nextStatus === 'Видано') {
      setNotice('Замовлення видано');
    }
  }

  function editSimpleManagerOrder(orderId: string, payload: { phone: string; client: string; device: string; issue: string; estimatedAmount: string; engineerId: string }) {
    const targetOrder = orders.find((order) => order.id === orderId);
    if (targetOrder && !ensureOrderEditable(targetOrder, 'Редагування замовлення')) return;
    const engineer = users.find((user) => user.id === payload.engineerId && user.role === 'Інженер');
    if (!engineer) {
      setNotice('Не вдалося зберегти: оберіть інженера.');
      return;
    }
    const normalizedPhone = payload.phone.trim();
    const trimmedClient = payload.client.trim();
    const trimmedDevice = payload.device.trim();
    const trimmedIssue = payload.issue.trim();
    const estimatedAmount = Number(payload.estimatedAmount);
    if (!normalizedPhone || !trimmedClient || !trimmedDevice || !trimmedIssue) {
      setNotice('Не вдалося зберегти: заповніть телефон, імʼя, пристрій та проблему.');
      return;
    }
    const clientExists = customerList.some((client) => client.phone.replace(/\D/g, '') === normalizedPhone.replace(/\D/g, ''));
    if (!clientExists) {
      setCustomerList((current) => [{ name: trimmedClient, phone: normalizedPhone, taxId: '', orders: 1 }, ...current]);
    } else {
      setCustomerList((current) => current.map((client) => (
        client.phone.replace(/\D/g, '') === normalizedPhone.replace(/\D/g, '')
          ? { ...client, name: trimmedClient }
          : client
      )));
    }
    persistOrdersUpdate((current) => current.map((item) => (
      item.id === orderId
        ? {
            ...item,
            phone: normalizedPhone,
            client: trimmedClient,
            device: trimmedDevice,
            issue: trimmedIssue,
            assignedTo: engineer.id,
            engineer: engineer.name,
            estimatedAmount: Number.isFinite(estimatedAmount) && estimatedAmount > 0 ? estimatedAmount : undefined,
          }
        : item
    )));
    appendSimpleOrderActivity(orderId, 'Замовлення оновлено', `${trimmedClient} · ${trimmedDevice}`);
    logAction('Редагування замовлення', orderId, `${trimmedClient} · ${trimmedDevice} · інженер ${engineer.name}.`);
    setNotice(`Замовлення ${orderId} оновлено.`);
  }

  function addSimpleManagerOrderPart(orderId: string, productId: string, qtyToAdd: number) {
    const requestedQty = Math.max(1, Math.floor(qtyToAdd));
    const targetOrder = orders.find((order) => order.id === orderId);
    const targetProduct = products.find((product) => product.id === productId);
    if (!targetOrder || !targetProduct) {
      setNotice('Не вдалося додати деталь.');
      return;
    }
    if (!ensureOrderEditable(targetOrder, 'Додавання деталі')) return;
    if (available(targetProduct) < requestedQty) {
      setNotice(`Недостатньо залишку для "${targetProduct.name}".`);
      return;
    }
    const fifo = fifoConsumeBatches(targetProduct, requestedQty);
    if (!fifo) {
      setNotice(`Не вдалося списати "${targetProduct.name}" по партіях FIFO.`);
      return;
    }
    const partPrice = targetProduct.price || targetProduct.cost;
    const newPart: OrderPart = {
      id: uid('PART'),
      productId: targetProduct.id,
      qty: requestedQty,
      status: 'Встановлено',
      cost: Math.round(fifo.totalCost / requestedQty),
      price: partPrice,
      batchAllocations: fifo.allocations,
    };
    const updatedProducts = products.map((product) => (
      product.id === targetProduct.id
        ? { ...product, stock: Math.max(product.stock - requestedQty, 0), installed: product.installed + requestedQty, batches: fifo.batches }
        : product
    ));
    const updatedOrders = orders.map((order) => (
      order.id === orderId
        ? {
            ...order,
            parts: [...order.parts, newPart],
            activityLog: [
              { id: uid('ACT'), date: today, action: 'Деталь додана', detail: `${targetProduct.name} · ${requestedQty}` },
              ...(order.activityLog ?? []),
            ],
          }
        : order
    ));
    saveProductsToStorage(updatedProducts);
    saveOrdersToStorage(updatedOrders);
    setProducts(updatedProducts);
    setOrders(updatedOrders);
    addMovement({
      type: 'Списання',
      productId: targetProduct.id,
      qty: requestedQty,
      orderId,
      basis: orderId,
      batchRefs: fifo.allocations.map((item) => `${item.batchId} x${item.qty}`).join(', '),
      unitPrice: requestedQty > 0 ? Math.round(fifo.totalCost / requestedQty) : 0,
      totalAmount: fifo.totalCost,
      comment: `Списано в ремонт по FIFO: ${fifo.allocations.map((item) => `${item.batchId} x${item.qty}`).join(', ')}.`,
    });
    logAction('Списання в ремонт', orderId, `${targetProduct.name} · ${requestedQty} шт. · FIFO ${fifo.allocations.map((item) => `${item.batchId}/${item.qty}`).join(', ')}.`);
    setNotice(`Деталь "${targetProduct.name}" додано до ${orderId}.`);
  }

  function removeSimpleManagerOrderPart(orderId: string, partId: string, reason?: string, comment?: string) {
    const targetOrder = orders.find((order) => order.id === orderId);
    const targetPart = targetOrder?.parts.find((part) => part.id === partId);
    if (!targetOrder || !targetPart) {
      setNotice('Не вдалося видалити деталь.');
      return;
    }
    if (!ensureOrderEditable(targetOrder, 'Повернення деталі')) return;
    const reasonText = reason?.trim() || 'Повернення деталі';
    const commentText = comment?.trim();
    const detailText = [reasonText, commentText].filter(Boolean).join(' · ');
    const updatedProducts = products.map((product) => (
      product.id === targetPart.productId
        ? {
            ...product,
            stock: product.stock + targetPart.qty,
            installed: Math.max(product.installed - targetPart.qty, 0),
            batches: targetPart.batchAllocations ? restoreBatchAllocations(product, targetPart.batchAllocations) : product.batches,
          }
        : product
    ));
    const updatedOrders = orders.map((order) => (
      order.id === orderId
        ? {
            ...order,
            parts: order.parts.filter((part) => part.id !== partId),
            activityLog: [
              { id: uid('ACT'), date: today, action: 'Повернення деталі', detail: `${products.find((product) => product.id === targetPart.productId)?.name ?? targetPart.productId} · ${targetPart.qty}${detailText ? ` · ${detailText}` : ''}` },
              ...(order.activityLog ?? []),
            ],
          }
        : order
    ));
    saveProductsToStorage(updatedProducts);
    saveOrdersToStorage(updatedOrders);
    setProducts(updatedProducts);
    setOrders(updatedOrders);
    addMovement({
      type: 'Повернення',
      productId: targetPart.productId,
      qty: targetPart.qty,
      orderId,
      basis: orderId,
      batchRefs: targetPart.batchAllocations?.map((item) => `${item.batchId} x${item.qty}`).join(', '),
      unitPrice: targetPart.cost,
      totalAmount: targetPart.cost * targetPart.qty,
      comment: `Деталь повернено з ремонту назад у складські партії.${detailText ? ` ${detailText}.` : ''}`,
    });
    logAction('Повернення деталі на склад', orderId, `${products.find((product) => product.id === targetPart.productId)?.name ?? targetPart.productId} · ${targetPart.qty} шт.${detailText ? ` · ${detailText}` : ''}`);
    setNotice(`Деталь повернуто на склад із ${orderId}.`);
  }

  function acceptSimpleManagerPayment(orderId: string, repairPriceInput: string, paymentType: 'наличные' | 'карта' | 'перевод', paymentKind: SimpleLedgerPaymentKind = 'оплата', paymentReason = '') {
    const targetOrder = orders.find((order) => order.id === orderId);
    if (!targetOrder) {
      setNotice('Замовлення не знайдено.');
      return null;
    }
    if (!ensureOrderEditable(targetOrder, 'Прийом оплати')) return null;
    if (targetOrder?.contractId || targetOrder?.legalEntity) {
      setNotice('Для юрособи оплата йде по акту, а не по замовленню.');
      return null;
    }
    const repairPrice = Number(repairPriceInput);
    if (!Number.isFinite(repairPrice) || repairPrice <= 0) {
      setNotice('Не можна прийняти оплату без ціни ремонту.');
      return null;
    }
    const orderTotal = Math.max(targetOrder.repairPrice ?? targetOrder.estimatedAmount ?? orderTotals(targetOrder).total, 0);
    const alreadyBooked = targetOrder.payments.filter(paymentCountsAsApplied).reduce((sum, payment) => sum + payment.amount, 0);
    const remaining = Math.max(orderTotal - alreadyBooked, 0);
    const acceptedPayment = Math.min(repairPrice, remaining);
    if (acceptedPayment <= 0) {
      setNotice('Оплата не потрібна: борг уже закрито.');
      return null;
    }
    if (!canAcceptShiftBasedPayment(paymentType)) {
      return null;
    }
    const ledgerStatus: SimpleLedgerPaymentStatus = paymentType === 'наличные'
      ? 'подтвержден'
      : paymentType === 'карта'
        ? 'проведено'
        : 'подтвержден';
    const orderPaymentStatus: Payment['status'] = paymentType === 'наличные'
      ? 'Підтверджено'
      : paymentType === 'карта'
        ? 'Проведено'
        : 'Підтверджено';
    const paymentRecord: SimpleOrderPaymentRecord = {
      id: uid('SMPAY'),
      client: targetOrder?.client ?? '',
      entityType: 'order',
      entityId: orderId,
      clientTaxId: clientTaxId(targetOrder.client, targetOrder.phone),
      orderId,
      amount: acceptedPayment,
      method: paymentType,
      paymentKind,
      status: ledgerStatus,
      date: today,
      acceptedBy: activeUser.name,
      cashShiftId: paymentType === 'перевод' ? undefined : cashShift.id,
      countedAtShiftId: paymentType === 'перевод' ? undefined : cashShift.id,
    };
    const orderPaymentRecord: Payment = {
      id: paymentRecord.id,
      date: today,
      amount: acceptedPayment,
      method: mapSimpleMethodToPaymentMethod(paymentType),
      type: mapSimpleKindToPaymentType(paymentKind),
      transactionNo: '',
      acceptedBy: activeUser.name,
      status: orderPaymentStatus,
      confirmedBy: orderPaymentStatus === 'Підтверджено' ? activeUser.name : undefined,
      confirmedAt: orderPaymentStatus === 'Підтверджено' ? today : undefined,
      orderId,
      comment: [paymentKind, paymentType, paymentReason.trim()].filter(Boolean).join(' · '),
      cashShiftId: paymentType === 'перевод' ? undefined : cashShift.id,
      countedAtShiftId: paymentType === 'перевод' ? undefined : cashShift.id,
    };
    const nextRemaining = Math.max(clientOrderDebt(targetOrder) - acceptedPayment, 0);
    const updatedOrders = orders.map((order) => (
      order.id === orderId
        ? {
            ...order,
            repairPrice: order.repairPrice ?? repairPrice,
            repairPaymentMethod: paymentType,
            status: nextRemaining <= 0
              ? 'Готовий до видачі'
              : ['Видано', 'Закрито', 'Скасовано'].includes(order.status)
                ? order.status
                : 'В ремонті',
            statusChangedAt: (
              (nextRemaining <= 0 ? 'Готовий до видачі' : ['Видано', 'Закрито', 'Скасовано'].includes(order.status) ? order.status : 'В ремонті') !== order.status
            ) ? today : order.statusChangedAt,
            vatStatus: order.legalEntity ? 'Очікує ПН' : order.vatStatus,
            payments: [
              orderPaymentRecord,
              ...order.payments,
            ],
            activityLog: [
              { id: uid('ACT'), date: today, action: 'Платіж', detail: `${acceptedPayment} · ${paymentKind} · ${paymentProcessingLabel(orderPaymentRecord)}${paymentReason.trim() ? ` · ${paymentReason.trim()}` : ''}` },
              ...(order.activityLog ?? []),
            ],
          }
        : order
    ));
    const updatedPayments = [paymentRecord, ...simplePayments];
    saveOrdersToStorage(updatedOrders);
    saveSimplePaymentsToStorage(updatedPayments);
    setOrders(updatedOrders);
    setSimplePayments(updatedPayments);
    if (paymentType !== 'перевод') {
      registerCashPayment(acceptedPayment, mapSimpleMethodToPaymentMethod(paymentType));
    }
    if (orderPaymentStatus === 'Підтверджено') {
      logAction('Підтвердження платежу', orderId, `${money(acceptedPayment)} підтверджено одразу (${paymentType}).`);
      if (nextRemaining <= 0) {
        appendSimpleOrderActivity(orderId, 'Борг закрито', `Платіж ${money(acceptedPayment)} підтверджено одразу`);
      }
    } else {
      logAction('Створення платежу', orderId, `${money(acceptedPayment)} · ${paymentKind} · ${paymentType} · ${paymentProcessingLabel(orderPaymentRecord)}.`);
    }
    setNotice(nextRemaining <= 0
      ? `Оплату для ${orderId} прийнято. Замовлення готове до видачі.`
      : `Оплату для ${orderId} прийнято. Залишок: ${money(nextRemaining)}.`);
    return paymentRecord.id;
  }

  function accountSimpleManagerOrderToContract(orderId: string) {
    const targetOrder = orders.find((order) => order.id === orderId);
    if (!targetOrder?.contractId) {
      setNotice('Для цього замовлення не вибрано договір.');
      return;
    }
    if (targetOrder.contractAccountedAt) {
      setNotice(`Замовлення ${orderId} уже враховано в договорі.`);
      return;
    }
    const targetContract = contracts.find((contract) => contract.id === targetOrder.contractId);
    if (!targetContract || targetContract.status !== 'Активний') {
      setNotice('Договір не знайдено або він уже закритий.');
      return;
    }
    const contractAmount = targetOrder.repairPrice ?? targetOrder.estimatedAmount ?? 0;
    if (!Number.isFinite(contractAmount) || contractAmount <= 0) {
      setNotice('Для врахування в договорі вкажіть суму ремонту.');
      return;
    }
    const usage = contractUsage(targetContract, orders);
    if (usage.used + contractAmount > targetContract.amount) {
      setNotice(`Ліміт договору перевищено. Доступно ще ${money(usage.remaining)}.`);
      return;
    }
    const updatedOrders = orders.map((order) => (
      order.id === orderId
        ? {
            ...order,
            contractAccountedAt: today,
            statusHistory: [
              { id: uid('H'), oldStatus: order.status, newStatus: order.status, changedBy: activeUser.name, changedAt: today, comment: `Суму ${money(contractAmount)} враховано в договорі ${targetContract.id}.` },
              ...(order.statusHistory ?? []),
            ],
            activityLog: [
              { id: uid('ACT'), date: today, action: 'Учтено в договор', detail: `${targetContract.id} · ${money(contractAmount)}` },
              ...(order.activityLog ?? []),
            ],
          }
        : order
    ));
    saveOrdersToStorage(updatedOrders);
    setOrders(updatedOrders);
    logAction('Врахування в договір', orderId, `${targetContract.id} · ${money(contractAmount)}.`);
    setNotice(`Замовлення ${orderId} враховано в договорі ${targetContract.id}.`);
  }

  function handleQuickPhoneChange(phone: string) {
    setQuickPhone(phone);
    const existingOrder = findExactOrderMatch(orders, phone);
    if (existingOrder) {
      setQuickClientDebtWarning(`Це номер існуючого замовлення. Відкрити замовлення ${existingOrder.id}?`);
      return;
    }
    const found = findClientBySearch(customerList, phone);
    if (found) {
      const debt = clientOutstandingDebtByPhone(found.phone);
      setQuickClientDebtWarning(debt > 0 ? `У клієнта є борг: ${money(debt)}` : '');
      setQuickClientName(found.name);
      return;
    }
    setQuickClientDebtWarning('');
  }

  function createQuickOrder() {
    const orderAuthor = viewUser;
    const engineer = users.find((user) => user.id === quickEngineerId && user.role === 'Інженер');
    const selectedContract = contracts.find((contract) => contract.id === quickContractId);
    const normalizedPhone = quickPhone.trim();
    const fallbackClientName = selectedContract?.client || quickClientName.trim() || `Клієнт ${normalizedPhone}`;
    if (!normalizedPhone || !quickDevice.trim() || !quickProblem.trim()) {
      setNotice('Створення замовлення: заповніть телефон, пристрій і проблему.');
      return;
    }
    const suggestedRepairLocation = quickLocationCode.trim() || suggestFreeLocation('REPAIR');
    const nextNumber = orders.reduce((max, order) => {
      const match = order.id.match(/(\d+)/);
      return Math.max(max, match ? Number(match[1]) : 0);
    }, 1040) + 1;
    const orderId = `ЗН-${nextNumber}`;
    const estimatedAmount = Number(quickEstimatedAmount);
    const clientExists = customerList.some((client) => client.phone.replace(/\D/g, '') === normalizedPhone.replace(/\D/g, ''));
    if (!clientExists) {
      setCustomerList((current) => [{ name: fallbackClientName, phone: normalizedPhone, taxId: '', orders: 1 }, ...current]);
    } else {
      setCustomerList((current) => current.map((client) => (client.phone.replace(/\D/g, '') === normalizedPhone.replace(/\D/g, '') ? { ...client, name: fallbackClientName, orders: client.orders + 1 } : client)));
    }

    const order: ServiceOrder = {
      id: orderId,
      groupId: `GR-${today.replace(/\D/g, '')}-${orders.length + 1}`,
      qrUrl: `/orders/${orderId}`,
      client: fallbackClientName,
      phone: normalizedPhone,
      device: quickDevice.trim(),
      serial: quickSerial.trim() || 'Не вказано',
      issue: quickProblem.trim(),
      appearance: quickAppearance.trim() || undefined,
      intakeComment: quickComment.trim() || undefined,
      estimatedAmount: Number.isFinite(estimatedAmount) && estimatedAmount > 0 ? estimatedAmount : undefined,
      contractId: selectedContract?.id,
      assignedTo: engineer?.id,
      createdByUserId: orderAuthor.id,
      engineer: engineer?.name ?? '',
      manager: orderAuthor.name,
      legalEntity: Boolean(selectedContract),
      status: 'Прийнято',
      createdAt: today,
      statusChangedAt: today,
      intakeDate: today,
      promisedDate: 'Уточнюється',
      deliveryAmount: 0,
      works: [],
      parts: [],
      payments: [],
      locationCode: suggestedRepairLocation,
      locationStatus: suggestedRepairLocation ? 'У комірці' : undefined,
      urgent: false,
      clientNotified: false,
      lastNotificationType: undefined,
      lastNotificationAt: undefined,
      notificationHistory: [],
      pendingExtraApproval: false,
      currentVersion: 1,
      statusHistory: [
        { id: uid('H'), newStatus: 'Прийнято', changedBy: orderAuthor.name, changedAt: today, comment: `Замовлення створено менеджером ${orderAuthor.name}${suggestedRepairLocation ? `, комірка ${suggestedRepairLocation}` : ''}.` },
      ],
      activityLog: [
        { id: uid('ACT'), date: today, action: 'Створено замовлення', detail: `${fallbackClientName} · ${quickDevice.trim()}` },
      ],
    };
    setOrders((current) => [order, ...current]);
    setOrderUnits((current) => [
      {
        id: `UNIT-${nextNumber}-1`,
        orderId,
        groupId: order.groupId,
        client: order.client,
        type: quickDevice.toLowerCase().includes('картридж') ? 'Картридж' : quickDevice.toLowerCase().includes('принтер') ? 'Принтер' : quickDevice.toLowerCase().includes('мфу') ? 'МФУ' : 'Ноутбук',
        code: `QR-UNIT-${nextNumber}-1`,
        status: 'Прийнято',
        engineer: engineer?.name ?? '',
        locationCode: suggestedRepairLocation,
        lastActionAt: today,
      },
      ...current,
    ]);
    if (suggestedRepairLocation) {
      setOrderMovementLogs((current) => [
        {
          id: uid('MOVE'),
          orderId,
          fromLocation: undefined,
          toLocation: suggestedRepairLocation,
          userId: orderAuthor.id,
          timestamp: today,
        },
        ...current,
      ]);
    }
    setOrderVersions((current) => [
      {
        id: uid('VER'),
        orderId,
        versionNo: 1,
        createdAt: today,
        createdBy: orderAuthor.name,
        reason: 'Початкове створення замовлення',
        snapshotData: JSON.stringify(order),
      },
      ...current,
    ]);
    setDocuments((current) => {
      const number = `НР-${current.length + 1001}`;
      const totals = documentVatTotals(0);
      return [{
        id: uid('DOC'),
        kind: 'Наряд на ремонт',
        number,
        entityType: 'service_order',
        entityId: orderId,
        clientOrSupplier: order.client,
        createdAt: today,
        createdBy: orderAuthor.name,
        version: 1,
        status: 'PDF збережено',
        pdfPath: `/documents/service_order/${orderId}/${number}.pdf`,
        amountNet: totals.amountNet,
        amountVat: totals.amountVat,
        amountGross: totals.amountGross,
      }, ...current];
    });
    setSelectedOrderId(orderId);
    enqueueClientNotification(order, 'Прийом пристрою');
    logAction('Швидкий прийом', orderId, `${order.client}: ${order.device}.${engineer ? ` Інженер: ${engineer.name}.` : ''}`);
    logAction('Створення замовлення', orderId, `${order.client} · ${order.device} · статус Прийнято.`);
    logAction('Автодокумент', orderId, 'Кнопка "Прийняти замовлення" автоматично створила наряд на ремонт.');
    setNotice(`${orderId}: замовлення прийнято, клієнта ${clientExists ? 'знайдено' : 'створено'}${engineer ? ', інженера призначено' : ''}.${suggestedRepairLocation ? ` Комірку ${suggestedRepairLocation} закріплено.` : ''} Наряд і наклейка готові до друку.`);
    setSuggestedDocumentAction({
      orderId,
      kind: 'Наряд на ремонт',
      title: 'Прийом завершено',
      text: 'Замовлення створено. Система одразу відкриває наряд без ручного заповнення.',
      autoOpen: true,
    });
    setQuickPhone('');
    setQuickClientDebtWarning('');
    setQuickClientName('');
    setQuickDevice('');
    setQuickSerial('');
    setQuickProblem('');
    setQuickAppearance('');
    setQuickEstimatedAmount('');
    setQuickEngineerId('');
    setQuickContractId('');
    setQuickComment('');
    setQuickLocationCode(suggestedRepairLocation ?? '');
  }

  function createContract(payload: { client: string; amount: string; startDate: string; endDate: string }) {
    const client = payload.client.trim();
    const amount = Number(payload.amount);
    if (!client || !payload.startDate || !payload.endDate || !Number.isFinite(amount) || amount <= 0) {
      setNotice('Заповніть клієнта, суму та строк договору.');
      return false;
    }
    const nextNumber = contracts.reduce((max, contract) => {
      const match = contract.id.match(/(\d+)/);
      return Math.max(max, match ? Number(match[1]) : 0);
    }, 1000) + 1;
    const nextContract: ContractRecord = {
      id: `DOG-${nextNumber}`,
      client,
      amount,
      startDate: payload.startDate,
      endDate: payload.endDate,
      status: 'Активний',
      createdAt: today,
    };
    setContracts((current) => [nextContract, ...current]);
    setNotice(`Договір ${nextContract.id} створено.`);
    return true;
  }

  function createContractFromOrders(payload: { client: string; startDate: string; endDate: string; orderIds: string[] }) {
    const client = payload.client.trim();
    const selectedOrders = orders.filter((order) => payload.orderIds.includes(order.id));
    if (!client || !payload.startDate || !payload.endDate || selectedOrders.length === 0) {
      setNotice('Оберіть замовлення та вкажіть строк договору.');
      return false;
    }
    if (selectedOrders.some((order) => order.contractId)) {
      setNotice('Деякі замовлення вже привʼязані до договору.');
      return false;
    }
    const amount = selectedOrders.reduce((sum, order) => sum + contractOrderAmount(order), 0);
    const nextNumber = contracts.reduce((max, contract) => {
      const match = contract.id.match(/(\d+)/);
      return Math.max(max, match ? Number(match[1]) : 0);
    }, 1000) + 1;
    const nextContract: ContractRecord = {
      id: `DOG-${nextNumber}`,
      client,
      amount,
      startDate: payload.startDate,
      endDate: payload.endDate,
      status: 'Активний',
      createdAt: today,
    };
    const updatedOrders = orders.map((order) => (
      payload.orderIds.includes(order.id)
        ? {
            ...order,
            contractId: nextContract.id,
            contractAccountedAt: today,
            legalEntity: true,
            activityLog: [
              { id: uid('ACT'), date: today, action: 'Додано в договір', detail: `${nextContract.id} · ${money(contractOrderAmount(order))}` },
              ...(order.activityLog ?? []),
            ],
          }
        : order
    ));
    saveOrdersToStorage(updatedOrders);
    setOrders(updatedOrders);
    setContracts((current) => [nextContract, ...current]);
    logAction('Договір із замовлень', nextContract.id, `${selectedOrders.length} замовлень додано в договір.`);
    setNotice(`Договір ${nextContract.id} створено з вибраних замовлень.`);
    return true;
  }

  function createContractAct(contractId: string, orderIds: string[]) {
    const contract = contracts.find((item) => item.id === contractId);
    if (!contract) {
      setNotice('Договір не знайдено.');
      return false;
    }
    const selectedOrders = orders.filter((order) => orderIds.includes(order.id) && order.contractId === contractId);
    if (selectedOrders.length === 0) {
      setNotice('Оберіть замовлення для акта.');
      return false;
    }
    if (selectedOrders.some((order) => !order.contractAccountedAt)) {
      setNotice('Усі замовлення в акті мають бути вже враховані в договорі.');
      return false;
    }
    if (selectedOrders.some((order) => order.contractActId)) {
      setNotice('Деякі замовлення вже закриті актом.');
      return false;
    }
    const nextNumber = contractActs.reduce((max, act) => {
      const match = act.id.match(/(\d+)/);
      return Math.max(max, match ? Number(match[1]) : 0);
    }, 1000) + 1;
    const amount = selectedOrders.reduce((sum, order) => sum + contractOrderAmount(order), 0);
    const nextAct: ContractActRecord = {
      id: `ACT-${nextNumber}`,
      contractId,
      orderIds: selectedOrders.map((order) => order.id),
      amount,
      paidAmount: 0,
      remainingAmount: amount,
      date: today,
      status: 'Не оплачено',
    };
    const updatedOrders = orders.map((order) => (
      orderIds.includes(order.id)
        ? {
            ...order,
            contractActId: nextAct.id,
            activityLog: [
              { id: uid('ACT'), date: today, action: 'Закрито актом', detail: `${nextAct.id} · ${money(amount)}` },
              ...(order.activityLog ?? []),
            ],
          }
        : order
    ));
    saveOrdersToStorage(updatedOrders);
    setOrders(updatedOrders);
    setContractActs((current) => [nextAct, ...current]);
    logAction('Акт договору', contractId, `${nextAct.id} створено на ${money(amount)}.`);
    setNotice(`Акт ${nextAct.id} створено.`);
    return true;
  }

  function createInvoiceForContractAct(actId: string) {
    const act = contractActs.find((item) => item.id === actId);
    if (!act) {
      setNotice('Акт не знайдено.');
      return;
    }
    const contract = contracts.find((item) => item.id === act.contractId);
    if (!contract) {
      setNotice('Договір для акта не знайдено.');
      return;
    }
    let created = false;
    setDocuments((current) => {
      const existing = current.find((document) => document.entityType === 'act' && document.entityId === act.id && document.kind === 'Рахунок на оплату');
      if (existing) return current;
      const number = nextDocumentNumber('Рахунок на оплату', current);
      const totals = invoiceTotals(act.amount, companySettings.vatEnabled);
      created = true;
      return [{
        id: uid('DOC'),
        kind: 'Рахунок на оплату',
        number,
        entityType: 'act',
        entityId: act.id,
        clientOrSupplier: contract.client,
        createdAt: today,
        createdBy: activeUser.name,
        version: 1,
        status: 'Очікує оплати',
        pdfPath: `/documents/act/${act.id}/${number}.pdf`,
        amountNet: totals.totalWithoutVAT,
        amountVat: totals.vat,
        amountGross: totals.totalWithVAT,
        snapshotData: JSON.stringify({
          actId: act.id,
          contractId: contract.id,
          client: contract.client,
          amount: act.amount,
          orders: orders
            .filter((order) => act.orderIds.includes(order.id))
            .map((order) => ({ id: order.id, device: order.device, amount: contractOrderAmount(order) })),
        }),
      }, ...current];
    });
    if (!created) {
      setNotice(`Рахунок для ${act.id} уже існує.`);
      return;
    }
    logAction('Рахунок по акту', act.id, `Рахунок сформовано на ${money(act.amount)}.`);
    setNotice(`Рахунок для ${act.id} сформовано.`);
  }

  function buildBankImportCandidates() {
    const orderCandidates: BankImportCandidate[] = orders
      .filter((order) => !order.contractId && clientOrderDebt(order) > 0)
      .map((order) => ({
        key: `order:${order.id}`,
        kind: 'order',
        entityId: order.id,
        client: order.client,
        amount: clientOrderDebt(order),
        label: `${order.id} · ${order.client}`,
        reference: order.id,
        taxId: clientTaxId(order.client, order.phone),
        eventDate: order.intakeDate,
      }));
    const actCandidates: BankImportCandidate[] = contractActs
      .filter((act) => act.remainingAmount > 0)
      .map((act) => {
        const contract = contracts.find((item) => item.id === act.contractId);
        return {
          key: `act:${act.id}`,
          kind: 'act',
          entityId: act.id,
          client: contract?.client ?? act.contractId,
          amount: act.remainingAmount,
          label: `${act.id} · ${contract?.client ?? act.contractId}`,
          reference: act.id,
          taxId: clientTaxId(contract?.client ?? act.contractId),
          eventDate: act.date,
        } satisfies BankImportCandidate;
      });
    const invoiceCandidates: BankImportCandidate[] = documents
      .filter((document) => document.kind === 'Рахунок на оплату' && document.status === 'Очікує оплати')
      .map((document) => ({
        key: `invoice:${document.id}`,
        kind: 'invoice',
        entityId: document.id,
        client: document.clientOrSupplier,
        amount: document.amountGross ?? 0,
        label: `${document.number} · ${document.clientOrSupplier}`,
        reference: document.number,
        taxId: clientTaxId(document.clientOrSupplier),
        eventDate: document.createdAt,
        backingEntityType: document.entityType,
        backingEntityId: document.entityId,
        documentId: document.id,
      }));
    const clientBalanceCandidates: BankImportCandidate[] = customerList.map((client) => ({
      key: `client:${client.phone || client.name}`,
      kind: 'invoice',
      entityId: client.phone || client.name,
      client: client.name,
      amount: 0,
      label: `${client.name} · загальний борг клієнта`,
      reference: client.name,
      taxId: normalizeTaxId(client.taxId),
      eventDate: today,
      backingEntityType: 'contract',
      backingEntityId: client.name,
    }));
    return [...orderCandidates, ...actCandidates, ...invoiceCandidates, ...clientBalanceCandidates];
  }

  function findDuplicateBankPayment(row: BankStatementRow) {
    const fingerprint = bankImportFingerprint(row);
    return simplePayments.find((payment) => payment.fingerprint === fingerprint);
  }

  function applyImportedBankPayment(row: BankStatementRow, candidate: BankImportCandidate) {
    if (findDuplicateBankPayment(row)) return false;
    const entityType: SimpleLedgerEntityType = candidate.kind === 'invoice'
      ? candidate.backingEntityType === 'act'
        ? 'act'
        : candidate.backingEntityType === 'service_order'
          ? 'order'
          : candidate.backingEntityId === candidate.client
            ? 'client'
            : 'order'
      : candidate.kind;
    const paymentRecord: SimpleOrderPaymentRecord = {
      id: uid('BANKPAY'),
      client: candidate.client,
      entityType,
      entityId: candidate.kind === 'invoice' ? (candidate.backingEntityId ?? candidate.entityId) : candidate.entityId,
      clientTaxId: candidate.taxId || normalizeTaxId(row.taxId),
      orderId: candidate.kind === 'order' ? candidate.entityId : candidate.backingEntityType === 'service_order' ? candidate.backingEntityId : undefined,
      actId: candidate.kind === 'act' ? candidate.entityId : candidate.backingEntityType === 'act' ? candidate.backingEntityId : undefined,
      amount: row.amount,
      method: 'перевод',
      paymentKind: 'оплата',
      status: 'подтвержден',
      date: row.date || today,
      currency: row.currency,
      payer: row.payer,
      taxId: normalizeTaxId(row.taxId),
      purpose: row.purpose,
      documentRef: row.documentRef,
      direction: row.direction,
      bankAccount: row.bankAccount,
      bankName: companyBankAccounts.find((account) => normalizePaymentText(account.iban) === normalizePaymentText(row.bankAccount))?.bankName ?? 'Невідомий банк',
      fingerprint: bankImportFingerprint(row),
    };

    if (candidate.kind === 'order' || candidate.backingEntityType === 'service_order') {
      const orderId = candidate.kind === 'order' ? candidate.entityId : candidate.backingEntityId;
      const targetOrder = orders.find((order) => order.id === orderId);
      if (!targetOrder) return false;
      const remaining = clientOrderDebt(targetOrder);
      if (row.amount - remaining > 0.01) return false;
      const orderPayment: Payment = {
        id: paymentRecord.id,
        date: row.date || today,
        amount: row.amount,
        method: 'Безготівка',
        type: row.amount >= remaining - 0.01 ? 'Повна оплата' : 'Часткова оплата',
        transactionNo: row.documentRef || uid('BANK'),
        acceptedBy: activeUser.name,
        status: 'Підтверджено',
        confirmedBy: activeUser.name,
        confirmedAt: today,
        orderId,
        comment: `Імпорт виписки: ${row.purpose || row.payer}`,
      };
      const updatedOrders = orders.map((order) => (
        order.id === orderId
          ? (() => {
              const nextRemaining = Math.max(remaining - row.amount, 0);
              const nextStatus = nextOrderStatusAfterAutoPayment(order, nextRemaining);
              return {
                ...order,
                status: nextStatus,
                statusChangedAt: nextStatus !== order.status ? today : order.statusChangedAt,
                payments: [orderPayment, ...order.payments],
                activityLog: [
                  { id: uid('ACT'), date: today, action: 'Банк підтвердив оплату', detail: `${money(row.amount)} · ${row.payer} · залишок ${money(nextRemaining)}` },
                  ...(order.activityLog ?? []),
                ],
              };
            })()
          : order
      ));
      saveOrdersToStorage(updatedOrders);
      setOrders(updatedOrders);
      logSystemAction('Автопідтвердження платежу', orderId ?? candidate.entityId, `${row.payer} · ${money(row.amount)} · точне співпадіння з замовленням.`);
    }

    if (candidate.kind === 'act' || candidate.backingEntityType === 'act') {
      const actId = candidate.kind === 'act' ? candidate.entityId : candidate.backingEntityId;
      if (!actId) return false;
      const targetAct = contractActs.find((act) => act.id === actId);
      if (!targetAct || row.amount - targetAct.remainingAmount > 0.01) return false;
      const nextPaymentsForActs = [paymentRecord, ...simplePayments];
      setContractActs(recalcContractActs(nextPaymentsForActs));
      setDocuments((current) => current.map((document) => (
        (document.id === candidate.documentId || (document.entityType === 'act' && document.entityId === actId && document.kind === 'Рахунок на оплату'))
          ? { ...document, status: row.amount >= targetAct.remainingAmount - 0.01 ? 'Оплачено' as PrintDocument['status'] : 'Очікує оплати' as PrintDocument['status'] }
          : document
      )));
      logSystemAction('Автопідтвердження платежу', actId, `${row.payer} · ${money(row.amount)} · точне співпадіння з актом.`);
    }

    if (entityType === 'client') {
      const clientDebtOrders = debtOrdersForClient({
        client: candidate.client,
        taxId: candidate.taxId || row.taxId,
      });
      if (clientDebtOrders.length > 0) {
        const shouldDistribute = clientDebtOrders.length === 1 ? true : window.confirm('Распределить оплату по долгам?');
        if (shouldDistribute) {
          const distribution = distributeClientDebtPayment(paymentRecord);
          if (distribution.allocations.length > 0) {
            saveOrdersToStorage(distribution.updatedOrders);
            setOrders(distribution.updatedOrders);
            logSystemAction(
              'Розподіл оплати по боргах',
              candidate.client,
              distribution.allocations.map((item) => `${item.orderId}: ${money(item.amount)} · залишок ${money(item.remainingAfter)}`).join('; '),
            );
            if (clientDebtOrders.length === 1) {
              setNotice(`Платіж автоматично розподілено на борг ${clientDebtOrders[0].id}.`);
            }
          }
        }
      }
    }

    if (candidate.kind === 'invoice' && candidate.backingEntityType !== 'act' && candidate.backingEntityId !== candidate.client && candidate.documentId) {
      setDocuments((current) => current.map((document) => (
        document.id === candidate.documentId
          ? { ...document, status: 'Оплачено' as PrintDocument['status'] }
          : document
      )));
    }

    const updatedPayments = [paymentRecord, ...simplePayments];
    saveSimplePaymentsToStorage(updatedPayments);
    setSimplePayments(updatedPayments);
    logAction('Імпорт виписки', candidate.entityId, `${row.payer} · ${money(row.amount)} · ${candidate.label}.`);
    return true;
  }

  function finalizeBankStatementImport(rows: Record<string, unknown>[], mapping: Partial<Record<BankImportField, string>>, fileLabel: string) {
    const parsedRows = importBankStatementRows(rows, mapping);
    const candidates = buildBankImportCandidates();
    const importedItems: BankImportItem[] = parsedRows.map((row) => {
      if (row.direction === 'outgoing') {
        return {
          id: uid('BANKITEM'),
          row,
          status: 'unmatched',
          reason: 'Вихідний платіж не може закривати документи',
          candidates: [],
        };
      }
      if (findDuplicateBankPayment(row)) {
        return {
          id: uid('BANKITEM'),
          row,
          status: 'unmatched',
          reason: 'Дублікат банківського платежу',
          candidates: [],
        };
      }
      const purposeText = normalizePaymentText(row.purpose);
      const payerText = normalizePaymentText(row.payer);
      const docText = normalizePaymentText(row.documentRef);
      const combinedText = [docText, purposeText].join(' ').trim();
      const rowTaxId = normalizeTaxId(row.taxId);
      const rowDateKey = normalizePaymentText(row.date);
      const referenceTokens = extractReferenceTokens(`${row.documentRef} ${row.purpose}`);
      const rank = (candidate: BankImportCandidate) => {
        const candidateRef = normalizePaymentText(candidate.reference);
        const candidateClient = normalizePaymentText(candidate.client);
        const candidateTaxId = normalizeTaxId(candidate.taxId);
        const refMatched = candidateRef && referenceTokens.some((token) => token === candidateRef || combinedText.includes(candidateRef));
        const sameDate = candidate.eventDate && normalizePaymentText(candidate.eventDate) === rowDateKey;
        if (refMatched && Math.abs(candidate.amount - row.amount) < 0.01) return 1;
        if (candidateTaxId && rowTaxId && candidateTaxId === rowTaxId) return 2;
        if (Math.abs(candidate.amount - row.amount) < 0.01 && payerText && (candidateClient.includes(payerText) || payerText.includes(candidateClient))) return 3;
        if (Math.abs(candidate.amount - row.amount) < 0.01 && sameDate) return 4;
        return 99;
      };
      const exact = candidates.filter((candidate) => rank(candidate) === 1).sort((a, b) => rank(a) - rank(b));
      if (exact.length === 1 && applyImportedBankPayment(row, exact[0])) {
        return {
          id: uid('BANKITEM'),
          row,
          status: 'matched',
          matchedCandidateKey: exact[0].key,
          matchedBy: 'auto',
          reason: 'Точний збіг',
          candidates: exact,
        };
      }
      const review = candidates.filter((candidate) => {
        const priority = rank(candidate);
        return priority <= 4;
      }).sort((a, b) => rank(a) - rank(b));
      if (review.length > 0) {
        return {
          id: uid('BANKITEM'),
          row,
          status: 'review',
          reason: review.length > 1 ? 'Знайдено кілька збігів, потрібна перевірка' : 'Потрібна ручна перевірка',
          candidates: review,
        };
      }
      return {
        id: uid('BANKITEM'),
        row,
        status: 'unmatched',
        reason: 'Платіж не розпізнано',
        candidates: [],
      };
    });
    setBankImportDraft(null);
    setBankImportItems(importedItems);
    logAction('Імпорт банківської виписки', fileLabel, `Рядків: ${importedItems.length}. Авто: ${importedItems.filter((item) => item.status === 'matched').length}. Перевірка: ${importedItems.filter((item) => item.status === 'review').length}. Нерозпізнані: ${importedItems.filter((item) => item.status === 'unmatched').length}.`);
    setNotice(`Виписку імпортовано: авто ${importedItems.filter((item) => item.status === 'matched').length}, перевірка ${importedItems.filter((item) => item.status === 'review').length}.`);
  }

  async function handleBankStatementImport(file: File) {
    const rows = await readSpreadsheetRows(file);
    const detection = detectBankImportMapping(rows);
    if (!detection.isRecognized) {
      setBankImportDraft({
        fileName: file.name,
        headers: detection.headers,
        rows,
        detectedBank: detection.detectedBank,
        mapping: detection.mapping,
      });
      logAction('Імпорт банківської виписки', file.name, 'Формат не розпізнано, очікується ручне зіставлення колонок.');
      setNotice('Формат CSV не розпізнано. Потрібно зіставити колонки вручну.');
      return;
    }
    finalizeBankStatementImport(rows, detection.mapping, file.name);
  }

  function finalizeMappedBankImport(mapping: Partial<Record<BankImportField, string>>) {
    if (!bankImportDraft) return;
    if (!mapping.date || !mapping.amount || !mapping.payer) {
      setNotice('Для імпорту потрібні колонки: дата, сума, платник.');
      return;
    }
    finalizeBankStatementImport(bankImportDraft.rows, mapping, bankImportDraft.fileName);
  }

  function confirmBankImportReview(itemId: string, candidateKey: string) {
    const item = bankImportItems.find((entry) => entry.id === itemId);
    const candidate = item?.candidates.find((entry) => entry.key === candidateKey) ?? buildBankImportCandidates().find((entry) => entry.key === candidateKey);
    if (!item || !candidate) {
      setNotice('Не вдалося підтвердити платіж із виписки.');
      return;
    }
    const applied = applyImportedBankPayment(item.row, candidate);
    if (!applied) {
      setNotice('Сума або документ не збігаються для безпечного підтвердження.');
      return;
    }
    setBankImportItems((current) => current.map((entry) => (
      entry.id === itemId
        ? { ...entry, status: 'matched', matchedCandidateKey: candidateKey, matchedBy: 'manual', reason: 'Підтверджено вручну' }
        : entry
    )));
    setNotice(`Платіж ${money(item.row.amount)} підтверджено вручну.`);
  }

  function closeContract(contractId: string) {
    const contract = contracts.find((item) => item.id === contractId);
    if (!contract) {
      setNotice('Договір не знайдено.');
      return;
    }
    if (contract.status === 'Закритий') {
      setNotice(`Договір ${contract.id} уже закритий.`);
      return;
    }
    setContracts((current) => current.map((item) => (
      item.id === contractId
        ? { ...item, status: 'Закритий', closedAt: today }
        : item
    )));
    logAction('Закриття договору', contract.id, `Договір закрито користувачем ${activeUser.name}.`);
    setNotice(`Договір ${contract.id} закрито.`);
  }

  function sendInternalMessage(toUserId: string, orderId: string, text: string, dueAt: string, importance: InternalMessageImportance) {
    const recipient = users.find((user) => user.id === toUserId);
    if (!recipient || !text.trim()) {
      setNotice('Повідомлення не створено: виберіть співробітника і напишіть задачу.');
      return;
    }
    const message: InternalMessage = {
      id: uid('MSG-IN'),
      toUser: recipient.name,
      toRole: recipient.role,
      orderId: orderId || undefined,
      text: text.trim(),
      dueAt: dueAt.trim() || today,
      importance,
      status: 'Нове',
      createdBy: activeUser.name,
      createdAt: today,
    };
    setInternalMessages((current) => [message, ...current]);
    logAction('Внутрішнє повідомлення', orderId || recipient.name, `Кому: ${recipient.role} — ${recipient.name}. ${message.text}`);
    setNotice(`Повідомлення для ${recipient.name} створено.`);
  }

  function updateInternalMessageStatus(messageId: string, status: InternalMessageStatus) {
    const message = internalMessages.find((item) => item.id === messageId);
    setInternalMessages((current) => current.map((item) => (item.id === messageId ? { ...item, status } : item)));
    if (message) {
      logAction('Статус повідомлення', message.orderId || message.toUser, `${message.toUser}: ${status}.`);
    }
  }

  function enqueueClientNotification(order: ServiceOrder, event: NotificationEvent, preferredChannel?: NotificationChannel, manual = false, messageOverride?: string) {
    const template = notificationTemplates.find((item) => item.event === event && item.enabled);
    const channel = preferredChannel ?? template?.channel ?? 'SMS';

    if (!order.phone.trim()) {
      const failed: ClientNotification = {
        id: uid('MSG'),
        orderId: order.id,
        client: order.client,
        phone: order.phone,
        channel,
        event,
        message: messageOverride ?? `Неможливо надіслати "${event}": у клієнта немає телефону.`,
        status: 'Помилка',
        createdAt: today,
        createdBy: activeUser.name,
        error: 'Не вказано номер телефону клієнта.',
      };
      setClientNotifications((current) => [failed, ...current]);
      persistOrdersUpdate((current) => current.map((item) => (
        item.id !== order.id
          ? item
          : {
              ...item,
              lastNotificationType: event,
              lastNotificationAt: today,
              notificationHistory: [failed, ...(item.notificationHistory ?? [])],
            }
      )));
      logAction('Помилка повідомлення клієнту', order.id, `${channel}: ${event}. Не вказано номер телефону.`);
      setNotice(`Повідомлення клієнту не надіслано: для ${order.id} не вказано телефон.`);
      return;
    }

    if (!template) {
      const disabled: ClientNotification = {
        id: uid('MSG'),
        orderId: order.id,
        client: order.client,
        phone: order.phone,
        channel,
        event,
        message: `Шаблон для події "${event}" вимкнено.`,
        status: 'Вимкнено',
        createdAt: today,
      };
      setClientNotifications((current) => [disabled, ...current]);
      setNotice(`Повідомлення не відправлено: шаблон "${event}" вимкнено.`);
      return;
    }

    const alreadySentToday = clientNotifications.some((item) => item.orderId === order.id && item.event === event && item.channel === channel && item.createdAt === today && item.status === 'Відправлено');
    if (alreadySentToday && !manual) {
      setNotice(`CRM не спамить клієнта: ${event} вже відправлено сьогодні.`);
      return;
    }

    const notification: ClientNotification = {
      id: uid('MSG'),
      orderId: order.id,
      client: order.client,
      phone: order.phone,
      channel,
      event,
      message: messageOverride ?? renderNotificationText(order, event),
      status: 'Відправлено',
      createdAt: today,
      createdBy: activeUser.name,
      templateName: notificationDisplay(event),
      sentAt: today,
    };
    setClientNotifications((current) => [notification, ...current]);
    setOrders((current) => current.map((item) => (
      item.id !== order.id
        ? item
        : {
            ...item,
            clientNotified: true,
            lastNotificationType: event,
            lastNotificationAt: today,
            notificationHistory: [notification, ...(item.notificationHistory ?? [])],
          }
    )));
    logAction(manual ? 'Повтор уведомлення' : 'Автоуведомлення клієнту', order.id, `${channel}: ${event}. ${notification.message}`);
    setNotice(`${channel}: клієнту ${order.client} відправлено повідомлення "${event}".`);
  }

  function createNotificationDraft(order: ServiceOrder, event: NotificationEvent, channel: NotificationChannel, text: string, status: NotificationStatus) {
    const notification: ClientNotification = {
      id: uid('MSG'),
      orderId: order.id,
      client: order.client,
      phone: order.phone,
      channel,
      event,
      message: text,
      status,
      createdAt: today,
      createdBy: activeUser.name,
      templateName: notificationDisplay(event),
      sentAt: status === 'Відправлено' ? today : undefined,
    };
    setClientNotifications((current) => [notification, ...current]);
    setOrders((current) => current.map((item) => (
      item.id !== order.id
        ? item
        : {
            ...item,
            clientNotified: status === 'Відправлено' ? true : item.clientNotified,
            lastNotificationType: event,
            lastNotificationAt: today,
            notificationHistory: [notification, ...(item.notificationHistory ?? [])],
          }
    )));
    return notification;
  }

  function notificationEventForStatus(status: OrderStatus): NotificationEvent | undefined {
    const map: Partial<Record<OrderStatus, NotificationEvent>> = {
      Прийнято: 'Прийом пристрою',
      'На діагностиці': 'Діагностика розпочата',
      'Очікує погодження': 'Діагностика завершена',
      Погоджено: 'Погодження прийнято',
      'Очікує запчастину': 'Очікування запчастини',
      'В ремонті': 'Ремонт розпочато',
      'На тестуванні': 'Ремонт завершено',
      'Готовий до видачі': 'Готово до видачі',
      'Очікує оплати': 'Очікує оплату',
      Видано: 'Видача',
      'Пауза / відкладено': 'Замовлення затримується',
    };
    return map[status];
  }

  function approvalClientUrl(token: string) {
    if (typeof window === 'undefined') return `/approval/${token}`;
    return `${window.location.origin}${window.location.pathname}#/approval/${token}`;
  }

  function sendRepairApproval(order: ServiceOrder, channel: NotificationChannel = 'Telegram', extra?: { description?: string; amount?: number; comment?: string }) {
    const totals = orderTotals(order);
    if (!order.diagnosisResult) {
      setNotice('Спочатку потрібно завершити діагностику і вказати результат.');
      return;
    }
    if (extra && (!extra.description?.trim() || !extra.amount || extra.amount <= 0)) {
      setNotice('Для погодження доп. робіт вкажіть опис і суму.');
      return;
    }
    if (!extra && totals.total <= 0) {
      setNotice('Согласування неможливе: немає суми робіт або запчастин.');
      return;
    }

    const token = uid('APPROVAL');
    const clientUrl = approvalClientUrl(token);
    const messageEvent: NotificationEvent = extra ? 'Потрібне погодження додаткових робіт' : 'Потрібне погодження додаткових робіт';
    const messageText = extra
      ? renderNotificationText(order, messageEvent, 3, { extraWorkDescription: extra.description, extraWorkAmount: extra.amount, comment: extra.comment })
      : `${order.client}, ремонт пристрою ${order.device}. Вартість: ${money(totals.total)}. Підтвердіть: ${clientUrl}`;
    const approval: RepairApproval = {
      id: uid('APR'),
      orderId: order.id,
      token,
      client: order.client,
      phone: order.phone,
      device: order.device,
      issue: order.issue,
      worksSnapshot: order.works.map((work) => ({ ...work })),
      partsSnapshot: order.parts.map((part) => ({ name: productName(part.productId), qty: part.qty, price: part.price })),
      totalAmount: extra ? totals.total + (extra.amount ?? 0) : totals.total,
      promisedDate: order.promisedDate,
      status: 'Очікує відповідь',
      sentAt: today,
      sentChannel: channel,
      messageText,
      extraWorkDescription: extra?.description?.trim(),
      extraWorkAmount: extra?.amount,
      managerComment: extra?.comment?.trim(),
      clientUrl,
    };

    setRepairApprovals((current) => [approval, ...current.filter((item) => !(item.orderId === order.id && item.status === 'Очікує відповідь'))]);
    setOrders((current) => current.map((item) => (
      item.id !== order.id
        ? item
        : {
            ...item,
            pendingExtraApproval: true,
            approvalComment: extra?.description ? `${extra.description}. ${extra.comment ?? ''}`.trim() : item.approvalComment,
          }
    )));
    updateOrderStatus(order.id, 'Очікує погодження', extra ? 'Менеджер відправив погодження додаткових робіт клієнту.' : 'Менеджер відправив клієнту сторінку погодження ремонту.');
    createNotificationDraft(order, messageEvent, channel, messageText, 'Відправлено');
    logAction('Відправка погодження', order.id, `${channel}: ${messageText}`);
    setNotice(`Погодження ${order.id} підготовлено. Менеджер може одразу скопіювати або відправити клієнту готовий текст.`);
  }

  function recordApprovalResponse(approvalId: string, accepted: boolean, manual = false) {
    const approval = repairApprovals.find((item) => item.id === approvalId);
    if (!approval) return;
    setRepairApprovals((current) =>
      current.map((item) =>
        item.id === approvalId
          ? {
              ...item,
              status: accepted ? 'Підтверджено' : 'Відмова',
              responseAt: today,
              responseIp: manual ? 'manual' : '93.77.120.14',
              responseComment: manual ? 'Менеджер зафіксував відповідь вручну.' : 'Клієнт відповів через сторінку погодження.',
            }
          : item,
      ),
    );
    setOrders((current) => current.map((item) => (
      item.id !== approval.orderId
        ? item
        : {
            ...item,
            pendingExtraApproval: false,
          }
    )));
    updateOrderStatus(approval.orderId, accepted ? 'Погоджено' : 'Пауза / відкладено', accepted ? 'Клієнт підтвердив ремонт через погодження.' : 'Клієнт відмовився або поставив ремонт на паузу.');
    const order = orders.find((item) => item.id === approval.orderId);
    if (order) {
      enqueueClientNotification(order, accepted ? 'Погодження прийнято' : 'Погодження відхилено', approval.sentChannel, true);
    }
    logAction(accepted ? 'Погодження ремонту' : 'Відмова від ремонту', approval.orderId, `${approval.client}: ${approval.totalAmount} грн. ${manual ? 'Вручну' : 'Через посилання'}.`);
    setNotice(accepted ? `${approval.orderId}: клієнт підтвердив ремонт.` : `${approval.orderId}: клієнт відмовився, заказ переведено в паузу.`);
  }

  function markApprovalNoAnswer(approvalId: string) {
    const approval = repairApprovals.find((item) => item.id === approvalId);
    if (!approval) return;
    setRepairApprovals((current) => current.map((item) => (item.id === approvalId ? { ...item, status: 'Немає відповіді' } : item)));
    const order = orders.find((item) => item.id === approval.orderId);
    if (order) enqueueClientNotification(order, 'Нагадування', 'SMS', true);
    logAction('Нагадування погодження', approval.orderId, 'Немає відповіді від клієнта, відправлено нагадування.');
    setNotice(`${approval.orderId}: відповіді немає, клієнту відправлено нагадування.`);
  }

  function createDocumentRecord(kind: DocumentKind, entityType: PrintDocument['entityType'], entityId: string, clientOrSupplier: string, status: PrintDocument['status'], sourceOrder?: ServiceOrder) {
    const sourceSale = sales.find((sale) => sale.id === entityId);
    if (kind === 'Видаткова накладна' && sourceOrder && sourceOrder.parts.length === 0) return undefined;
    if (kind === 'Видаткова накладна' && sourceSale && sourceSale.items.length === 0) return undefined;
    const gross = sourceOrder ? orderTotals(sourceOrder).total : sourceSale ? saleTotals(sourceSale).total : 0;
    const totals = documentVatTotals(gross);
    const number = `${documentPrefix[kind]}-${documents.length + 1001}`;
    return {
      id: uid('DOC'),
      kind,
      number,
      entityType,
      entityId,
      clientOrSupplier,
      createdAt: today,
      createdBy: activeUser.name,
      version: 1,
      status,
      pdfPath: `/documents/${entityType}/${entityId}/${number}.pdf`,
      amountNet: totals.amountNet,
      amountVat: totals.amountVat,
      amountGross: totals.amountGross,
    } satisfies PrintDocument;
  }

  function findServiceOrderDocument(orderId: string, kind: DocumentKind) {
    return documents.find((document) => document.entityType === 'service_order' && document.entityId === orderId && document.kind === kind);
  }

  function createServiceOrderDocument(kind: ServiceOrderDocumentKind, order: ServiceOrder) {
    const existing = findServiceOrderDocument(order.id, kind);
    if (existing) {
      setNotice(`${kind} вже створено: ${existing.number}.`);
      return existing;
    }
    let created: PrintDocument | undefined;
    setDocuments((current) => {
      const currentExisting = current.find((document) => document.entityType === 'service_order' && document.entityId === order.id && document.kind === kind);
      if (currentExisting) {
        created = currentExisting;
        return current;
      }
      const invoiceDocument = current.find((document) => document.entityType === 'service_order' && document.entityId === order.id && document.kind === 'Рахунок на оплату');
      if (kind === 'Акт надання послуг' && !invoiceDocument) return current;
      if (kind === 'Видаткова накладна' && order.parts.filter((part) => part.status !== 'Повернення').length === 0) return current;
      const number = nextDocumentNumber(kind, current);
      const latestTaxInvoice = findLatestOrderTaxInvoice(taxInvoices, order.id);
      const latestAppliedPayment = order.payments.find((payment) => paymentCountsAsApplied(payment));
      const snapshot = buildServiceOrderDocumentSnapshot(
        kind,
        order,
        products,
        number,
        companySettings,
        clientTaxId(order.client, order.phone),
        {
          invoiceNumber: invoiceDocument?.number,
          invoiceDate: invoiceDocument?.createdAt,
          invoiceId: invoiceDocument?.id,
          taxInvoiceId: latestTaxInvoice?.id,
          paymentId: latestAppliedPayment?.id,
        },
      );
      const vatTotals = invoiceTotals(orderTotals(order).total, companySettings.vatEnabled);
      const status: PrintDocument['status'] = kind === 'Акт надання послуг'
        ? 'Створено'
        : kind === 'Видаткова накладна'
          ? 'Створено'
          : kind === 'Рахунок на оплату'
            ? (orderDebtAmount(order) > 0 ? 'Очікує оплати' : 'Готово')
            : kind === 'Гарантійний талон'
              ? 'Створено'
              : 'Чернетка';
      const documentId = uid('DOC');
      created = {
        id: documentId,
        kind,
        number,
        entityType: 'service_order',
        entityId: order.id,
        orderId: order.id,
        clientOrSupplier: order.client,
        createdAt: today,
        createdBy: activeUser.name,
        version: 1,
        status,
        pdfPath: `/documents/service_order/${order.id}/${number}.pdf`,
        amountNet: vatTotals.totalWithoutVAT,
        amountVat: vatTotals.vat,
        amountGross: vatTotals.totalWithVAT,
        paymentId: latestAppliedPayment?.id,
        invoiceId: kind === 'Рахунок на оплату' ? documentId : invoiceDocument?.id,
        actId: kind === 'Акт надання послуг' ? documentId : current.find((document) => document.entityType === 'service_order' && document.entityId === order.id && document.kind === 'Акт надання послуг')?.id,
        taxInvoiceId: latestTaxInvoice?.id,
        waybillId: kind === 'Видаткова накладна' ? documentId : current.find((document) => document.entityType === 'service_order' && document.entityId === order.id && document.kind === 'Видаткова накладна')?.id,
        snapshotData: JSON.stringify(snapshot),
        orderVersionNo: order.currentVersion ?? 1,
      };
      return [created, ...current];
    });
    if (created) {
      logAction('Створення документа', order.id, `${kind} ${created.number}: створено.`);
      setNotice(`${kind} ${created.number} створено.`);
    } else if (kind === 'Акт надання послуг') {
      setNotice('Акт не можна створити без рахунку. Спочатку сформуйте рахунок.');
    } else if (kind === 'Видаткова накладна') {
      setNotice('Видаткову накладну можна створити тільки якщо в замовленні є товарні позиції.');
    }
    return created;
  }

  function printServiceOrderDocument(kind: ServiceOrderDocumentKind, order: ServiceOrder) {
    const existing = findServiceOrderDocument(order.id, kind);
    if (!existing) {
      setNotice(`Спочатку створіть документ "${kind}".`);
      return;
    }
    if (existing.orderVersionNo && existing.orderVersionNo !== (order.currentVersion ?? 1)) {
      const confirmed = typeof window !== 'undefined'
        ? window.confirm(`${orderDocumentsMismatchText(order)} Повторно відкрити старий документ ${existing.number}?`)
        : true;
      if (!confirmed) {
        logAction('Блокування повторного друку', order.id, `${kind} ${existing.number}: дані замовлення змінені після створення документа.`);
        setNotice(orderDocumentsMismatchText(order));
        return;
      }
    }
    markDocumentPrinted(order, kind);
    logAction(existing.printedAt ? 'Повторний друк документа' : 'Друк документа', order.id, `${kind} ${existing.number}: відкрито на друк.`);
    setNotice(`Документ "${kind}" ${existing.number} підготовлено до друку.`);
  }

  function addDocumentIfMissing(kind: DocumentKind, entityType: PrintDocument['entityType'], entityId: string, clientOrSupplier: string, status: PrintDocument['status'], sourceOrder?: ServiceOrder) {
    let created: PrintDocument | undefined;
    setDocuments((current) => {
      const existing = current.find((document) => document.kind === kind && document.entityId === entityId);
      if (existing) return current;
      const sourceSale = sales.find((sale) => sale.id === entityId);
      if (kind === 'Акт надання послуг' && sourceOrder && !current.find((document) => document.entityType === 'service_order' && document.entityId === sourceOrder.id && document.kind === 'Рахунок на оплату')) return current;
      if (kind === 'Видаткова накладна' && sourceOrder && sourceOrder.parts.length === 0) return current;
      if (kind === 'Видаткова накладна' && sourceSale && sourceSale.items.length === 0) return current;
      const gross = sourceOrder ? orderTotals(sourceOrder).total : sourceSale ? saleTotals(sourceSale).total : 0;
      const vatTotals = invoiceTotals(gross, companySettings.vatEnabled);
      const number = nextDocumentNumber(kind, current);
      const invoiceDocument = sourceOrder ? current.find((document) => document.entityType === 'service_order' && document.entityId === sourceOrder.id && document.kind === 'Рахунок на оплату') : undefined;
      const linkedActDocument = sourceOrder ? current.find((document) => document.entityType === 'service_order' && document.entityId === sourceOrder.id && document.kind === 'Акт надання послуг') : undefined;
      const linkedTaxInvoice = sourceOrder ? findLatestOrderTaxInvoice(taxInvoices, sourceOrder.id) : undefined;
      const linkedWaybillDocument = sourceOrder ? current.find((document) => document.entityType === 'service_order' && document.entityId === sourceOrder.id && document.kind === 'Видаткова накладна') : undefined;
      const latestAppliedPayment = sourceOrder?.payments.find((payment) => paymentCountsAsApplied(payment));
      const snapshotData = sourceOrder && ['Рахунок на оплату', 'Акт надання послуг', 'Видаткова накладна', 'Гарантійний талон'].includes(kind)
        ? JSON.stringify(buildServiceOrderDocumentSnapshot(kind as ServiceOrderDocumentKind, sourceOrder, products, number, companySettings, clientTaxId(sourceOrder.client, sourceOrder.phone), {
          invoiceNumber: invoiceDocument?.number,
          invoiceDate: invoiceDocument?.createdAt,
          invoiceId: invoiceDocument?.id,
          actId: linkedActDocument?.id,
          taxInvoiceId: linkedTaxInvoice?.id,
          waybillId: linkedWaybillDocument?.id,
          paymentId: latestAppliedPayment?.id,
        }))
        : undefined;
      const documentId = uid('DOC');
      created = {
        id: documentId,
        kind,
        number,
        orderId: sourceOrder?.id,
        entityType,
        entityId,
        clientOrSupplier,
        createdAt: today,
        createdBy: activeUser.name,
        version: 1,
        status,
        pdfPath: `/documents/${entityType}/${entityId}/${number}.pdf`,
        amountNet: vatTotals.totalWithoutVAT,
        amountVat: vatTotals.vat,
        amountGross: vatTotals.totalWithVAT,
        paymentId: latestAppliedPayment?.id,
        invoiceId: kind === 'Рахунок на оплату' ? documentId : invoiceDocument?.id,
        actId: kind === 'Акт надання послуг' ? documentId : linkedActDocument?.id,
        taxInvoiceId: linkedTaxInvoice?.id,
        waybillId: kind === 'Видаткова накладна' ? documentId : linkedWaybillDocument?.id,
        snapshotData,
        orderVersionNo: sourceOrder?.currentVersion ?? 1,
      };
      return [created, ...current];
    });
    if (created) logAction('Автодокумент', entityId, `${kind} ${created.number}: статус "${status}".`);
    return created;
  }

  function ensureOrderDocumentRecord(kind: 'Рахунок на оплату' | 'Акт надання послуг' | 'Наряд на ремонт' | 'Гарантійний талон', order: ServiceOrder) {
    const statusByKind: Record<typeof kind, PrintDocument['status']> = {
      'Наряд на ремонт': 'Чернетка',
      'Акт надання послуг': 'Створено',
      'Рахунок на оплату': orderTotals(order).debt > 0 ? 'Очікує оплати' : 'Готово',
      'Гарантійний талон': 'Створено',
    };
    return addDocumentIfMissing(kind, 'service_order', order.id, order.client, statusByKind[kind], order);
  }

  function markDocumentPrinted(order: ServiceOrder, kind: ServiceOrderDocumentKind) {
    setDocuments((current) => current.map((document) => (
      document.entityType === 'service_order' && document.entityId === order.id && document.kind === kind
        ? {
            ...document,
            status: document.kind === 'Акт надання послуг' || document.kind === 'Гарантійний талон' || document.kind === 'Видаткова накладна' ? 'Роздруковано' : document.status,
            printedAt: document.printedAt ?? today,
            printedBy: document.printedBy ?? activeUser.name,
          }
        : document
    )));
  }

  function signOrderAct(order: ServiceOrder) {
    let signed = false;
    setDocuments((current) => current.map((document) => {
      if (document.entityType === 'service_order' && document.entityId === order.id && document.kind === 'Акт надання послуг') {
        signed = true;
        return { ...document, status: 'Підписано', signedAt: today, signedBy: activeUser.name };
      }
      return document;
    }));
    if (!signed) {
      setNotice('Неможливо підписати акт: акт ще не створено.');
      return;
    }
    logAction('Акт підписано', order.id, `Акт підписано користувачем ${activeUser.name}.`);
    setOrders((current) => current.map((item) => (
      item.id !== order.id
        ? item
        : {
            ...item,
            statusHistory: [
              { id: uid('H'), oldStatus: item.status, newStatus: item.status, changedBy: activeUser.name, changedAt: today, comment: 'Акт виконаних робіт підписано.' },
              ...item.statusHistory,
            ],
          }
    )));
    setNotice(`Акт для ${order.id} підписано. Замовлення можна завершувати далі по процесу.`);
  }

  function logOrderDocumentPrint(order: ServiceOrder, kind: 'Рахунок на оплату' | 'Акт надання послуг' | 'Наряд на ремонт' | 'Гарантійний талон') {
    markDocumentPrinted(order, kind);
    setOrders((current) => current.map((item) => (
      item.id !== order.id
        ? item
        : {
            ...item,
            statusHistory: [
              {
                id: uid('H'),
                oldStatus: item.status,
                newStatus: item.status,
                changedBy: activeUser.name,
                changedAt: today,
                comment: `Роздруковано документ "${kind}". Користувач: ${activeUser.name}.`,
              },
              ...item.statusHistory,
            ],
          }
    )));
    logAction('Друк документа', order.id, `${kind}: роздруковано ${activeUser.name} о ${today}.`);
  }

  function latestTaxInvoiceForOrder(orderId: string) {
    return [...taxInvoices]
      .filter((invoice) => invoice.orderId === orderId)
      .sort((a, b) => (parseDateTime(b.createdAt)?.getTime() ?? 0) - (parseDateTime(a.createdAt)?.getTime() ?? 0))[0];
  }

  function taxInvoiceEventForOrder(order: ServiceOrder) {
    const signedAct = documents.find((document) => (
      document.entityType === 'service_order'
      && document.entityId === order.id
      && document.kind === 'Акт надання послуг'
      && document.status === 'Підписано'
    ));
    const actEvent = signedAct
      ? {
          type: 'act' as const,
          eventDate: signedAct.signedAt ?? signedAct.printedAt ?? today,
          actId: signedAct.id,
          actNumber: signedAct.number,
        }
      : null;
    const paymentEvent = [...order.payments]
      .filter((payment) => (
        payment.amount > 0
        && (
          (payment.method === 'Готівка' && payment.status === 'Підтверджено')
          || (payment.method === 'Картка' && payment.status === 'Проведено')
          || (payment.method === 'Безготівка' && payment.status === 'Підтверджено')
        )
      ))
      .sort((a, b) => (parseDateTime(a.confirmedAt ?? a.date)?.getTime() ?? 0) - (parseDateTime(b.confirmedAt ?? b.date)?.getTime() ?? 0))[0];
    const paymentCandidate = paymentEvent
      ? {
          type: 'payment' as const,
          eventDate: paymentEvent.confirmedAt ?? paymentEvent.date,
          paymentId: paymentEvent.id,
          paymentType: paymentEvent.method,
          paymentStatus: paymentEvent.status ?? 'Підтверджено',
        }
      : null;
    if (actEvent && paymentCandidate) {
      const actTime = parseDateTime(actEvent.eventDate)?.getTime() ?? 0;
      const paymentTime = parseDateTime(paymentCandidate.eventDate)?.getTime() ?? 0;
      return paymentTime > 0 && paymentTime <= actTime ? paymentCandidate : actEvent;
    }
    return paymentCandidate ?? actEvent;
  }

  function buildTaxInvoiceSnapshot(order: ServiceOrder, documentNumber: string, event: NonNullable<ReturnType<typeof taxInvoiceEventForOrder>>): TaxInvoiceSnapshot {
    const clientTaxCode = clientTaxId(order.client, order.phone);
    const totals = orderTotals(order);
    const vatTotals = invoiceTotals(totals.total, true);
    const invoiceDocument = findServiceOrderDocument(order.id, 'Рахунок на оплату');
    const waybillDocument = findServiceOrderDocument(order.id, 'Видаткова накладна');
    return {
      companyName: companySettings.companyName,
      companyEdrpou: companySettings.edrpou,
      companyIpn: companySettings.ipn,
      companyIban: companySettings.iban,
      companyBank: companySettings.bank,
      companyMfo: companySettings.mfo,
      companyAddress: companySettings.address,
      companyPhone: companySettings.phone,
      companyVatStatus: companySettings.vatEnabled ? 'Платник ПДВ' : 'Без ПДВ',
      clientName: order.client,
      clientPhone: order.phone,
      clientTaxId: clientTaxCode,
      orderId: order.id,
      orderNumber: order.id,
      device: order.device,
      managerName: order.manager,
      documentNumber,
      documentDate: extractDayKey(event.eventDate) || today,
      eventDate: extractDayKey(event.eventDate) || today,
      eventType: event.type,
      sumWithoutVAT: Math.round(vatTotals.totalWithoutVAT * 100) / 100,
      vatAmount: Math.round(vatTotals.vat * 100) / 100,
      sumWithVAT: Math.round(vatTotals.totalWithVAT * 100) / 100,
      totalWithVATWords: numberToWordsUA(vatTotals.totalWithVAT),
      vatAmountWords: numberToWordsUA(vatTotals.vat),
      lines: serviceOrderTaxInvoiceLines(order, products),
      paymentId: event.type === 'payment' ? event.paymentId : undefined,
      paymentType: event.type === 'payment' ? event.paymentType : undefined,
      paymentStatus: event.type === 'payment' ? event.paymentStatus : undefined,
      invoiceId: invoiceDocument?.id,
      actId: event.type === 'act' ? event.actId : undefined,
      actNumber: event.type === 'act' ? event.actNumber : undefined,
      waybillId: waybillDocument?.id,
      createdAt: today,
    };
  }

  function createTaxInvoiceForOrder(orderId: string) {
    const order = orders.find((item) => item.id === orderId);
    if (!order) {
      setNotice('Замовлення для ПН не знайдено.');
      return;
    }
    if (!order.legalEntity && !order.vatStatus) {
      setNotice('Податкова накладна доступна тільки для замовлень з ПДВ.');
      return;
    }
    const invoiceDocument = findServiceOrderDocument(order.id, 'Рахунок на оплату');
    if (!invoiceDocument) {
      logAction('Помилка ПН', orderId, 'ПН не створено: спочатку потрібно сформувати рахунок.');
      setNotice('Спочатку потрібно створити рахунок, і лише потім ПН.');
      return;
    }
    const event = taxInvoiceEventForOrder(order);
    if (!event) {
      logAction('Помилка ПН', orderId, 'ПН не створено: немає першої події (підписаного акта або підтвердженої оплати).');
      setNotice('ПН можна створити тільки після підписання акта або підтвердженої оплати.');
      return;
    }
    const clientTaxCode = clientTaxId(order.client, order.phone);
    if (!clientTaxCode) {
      logAction('Помилка ПН', orderId, 'ПН не створено: у клієнта відсутній ЄДРПОУ / ІПН.');
      setNotice('Для створення ПН потрібно вказати ЄДРПОУ / ІПН клієнта.');
      return;
    }
    const existing = taxInvoices.find((invoice) => (
      invoice.orderId === orderId
      && (
        (event.type === 'payment' && invoice.paymentId === event.paymentId)
        || (event.type === 'act' && invoice.actId === event.actId)
      )
    ));
    if (existing) {
      setOrders((current) => current.map((item) => (item.id === orderId ? { ...item, vatStatus: existing.status === 'Зареєстровано' ? 'ПН зареєстрована' : 'ПН створено' } : item)));
      setNotice(`Податкова накладна вже створена: ${existing.number}.`);
      return;
    }
    const number = nextTaxInvoiceNumber(taxInvoices);
    const snapshot = buildTaxInvoiceSnapshot(order, number, event);
    const invoice: TaxInvoice = {
      id: uid('TIN'),
      number,
      date: snapshot.documentDate,
      orderId,
      invoiceId: snapshot.invoiceId,
      actId: snapshot.actId,
      paymentId: snapshot.paymentId,
      waybillId: snapshot.waybillId,
      eventType: snapshot.eventType,
      client: order.client,
      eventDate: snapshot.eventDate,
      amount: snapshot.sumWithVAT,
      vatAmount: snapshot.vatAmount,
      status: 'Створено',
      snapshot: JSON.stringify(snapshot),
      createdAt: today,
      responsible: activeUser.name,
    };
    setTaxInvoices((current) => {
      const updated = [invoice, ...current];
      saveTaxInvoicesToStorage(updated);
      return updated;
    });
    setDocuments((current) => current.map((document) => (
      document.entityType === 'service_order'
      && document.entityId === orderId
      && ['Рахунок на оплату', 'Акт надання послуг', 'Видаткова накладна'].includes(document.kind)
        ? { ...document, taxInvoiceId: invoice.id }
        : document
    )));
    setOrders((current) => current.map((item) => (item.id === orderId ? { ...item, vatStatus: 'ПН створено' } : item)));
    logAction('Створення ПН', orderId, `${invoice.number} створено по події "${event.type === 'payment' ? 'payment' : 'act'}" на ${money(invoice.amount)}. Відповідальний: ${activeUser.name}.`);
    setNotice(`Податкову накладну ${invoice.number} створено.`);
  }

  function registerTaxInvoice(orderId: string) {
    const invoice = latestTaxInvoiceForOrder(orderId);
    if (!invoice) {
      setNotice('Спочатку потрібно створити податкову накладну.');
      return;
    }
    if (invoice.status === 'Зареєстровано') {
      setNotice(`Податкова накладна ${invoice.number} уже зареєстрована.`);
      return;
    }
    setTaxInvoices((current) => {
      const updated = current.map((item) => (
        item.id === invoice.id
          ? { ...item, status: 'Зареєстровано' as TaxInvoiceStatus, registeredAt: today }
          : item
      ));
      saveTaxInvoicesToStorage(updated);
      return updated;
    });
    setOrders((current) => current.map((item) => (item.id === orderId ? { ...item, vatStatus: 'ПН зареєстрована' } : item)));
    logAction('Реєстрація ПН', orderId, `${invoice.number} позначено як зареєстровану ${activeUser.name}.`);
    setNotice(`Податкову накладну ${invoice.number} позначено як зареєстровану.`);
  }

  function createSmartReadyDocuments(order: ServiceOrder) {
    const createdKinds: string[] = [];
    const act = addDocumentIfMissing('Акт надання послуг', 'service_order', order.id, order.client, 'Створено', order);
    if (act) {
      createdKinds.push('акт');
      logAction('Створено акт', order.id, `Створено ${act.number} для замовлення ${order.id}.`);
    }
    const invoice = addDocumentIfMissing('Рахунок на оплату', 'service_order', order.id, order.client, 'Очікує оплати', order);
    if (invoice) createdKinds.push('рахунок');
    if (order.parts.length > 0) {
      const delivery = addDocumentIfMissing('Видаткова накладна', 'service_order', order.id, order.client, 'Готово до видачі', order);
      if (delivery) createdKinds.push('видаткова');
    }
    setOrders((current) => current.map((item) => (item.id === order.id ? { ...item, vatStatus: 'Очікує ПН', actIssuedAt: item.actIssuedAt ?? today, debtSince: item.debtSince ?? today } : item)));
    if (orderTotals(order).debt > 0 || daysSince(order.statusChangedAt) > 2) {
      setInternalMessages((current) => [
        {
          id: uid('MSG-IN'),
          toUser: order.manager,
          toRole: 'Менеджер',
          orderId: order.id,
          text: 'Акт сформовано, але оплати ще немає. Потрібно проконтролювати оплату клієнта.',
          dueAt: today,
          importance: 'Важлива',
          status: 'Нове',
          createdBy: 'Система',
          createdAt: today,
        },
        ...(daysSince(order.statusChangedAt) > 2 ? [{
          id: uid('MSG-IN'),
          toUser: order.manager,
          toRole: 'Менеджер' as Role,
          orderId: order.id,
          text: `Заказ без руху ${daysSince(order.statusChangedAt)} дн. після статусу "${order.status}". Потрібен контроль.`,
          dueAt: today,
          importance: 'Важлива' as InternalMessageImportance,
          status: 'Нове' as InternalMessageStatus,
          createdBy: 'Система',
          createdAt: today,
        }] : []),
        {
          id: uid('MSG-IN'),
          toUser: 'Бухгалтер',
          toRole: 'Бухгалтер',
          orderId: order.id,
          text: 'Акт сформовано, підпис клієнта ще не зафіксовано. НДС: готово до формування даних для НН.',
          dueAt: today,
          importance: 'Важлива',
          status: 'Нове',
          createdBy: 'Система',
          createdAt: today,
        },
        ...current,
      ]);
    }
    logAction('Автоматика документів', order.id, `Готово: створено ${createdKinds.length ? createdKinds.join(', ') : 'без нових документів, дублікати вже існують'}. НДС: готово до формування даних для НН.`);
    return createdKinds;
  }

  function printDocument(kind: DocumentKind, entityType: PrintDocument['entityType'], entityId: string, clientOrSupplier: string) {
    const sourceOrder = orders.find((order) => order.id === entityId);
    const sourceSale = sales.find((sale) => sale.id === entityId);
    const existing = documents.find((document) => document.kind === kind && document.entityId === entityId);
    if (existing) {
      if (sourceOrder && existing.orderVersionNo && existing.orderVersionNo !== (sourceOrder.currentVersion ?? 1)) {
        const confirmed = typeof window !== 'undefined'
          ? window.confirm(`${orderDocumentsMismatchText(sourceOrder)} Повторно відкрити старий документ ${existing.number}?`)
          : true;
        if (!confirmed) {
          logAction('Блокування повторного друку', entityId, `${kind} ${existing.number}: дані замовлення змінені після створення документа.`);
          setNotice(orderDocumentsMismatchText(sourceOrder));
          return;
        }
      }
      if (sourceOrder && ['Рахунок на оплату', 'Акт надання послуг', 'Наряд на ремонт', 'Гарантійний талон'].includes(kind)) {
        logOrderDocumentPrint(sourceOrder, kind as 'Рахунок на оплату' | 'Акт надання послуг' | 'Наряд на ремонт' | 'Гарантійний талон');
      } else {
        logAction('Друк документа', entityId, `${kind} повторно відкрито для друку.`);
      }
      setNotice(`Документ "${kind}" вже є в реєстрі: ${existing.number}. Відкриваємо друк.`);
      window.print();
      return;
    }

    const number = nextDocumentNumber(kind, documents);
    if (kind === 'Видаткова накладна' && sourceOrder && sourceOrder.parts.length === 0) {
      setNotice('Видаткова накладна створюється тільки якщо в документі є товар або комплектуючі зі складу.');
      return;
    }
    if (kind === 'Акт надання послуг' && sourceOrder && !findServiceOrderDocument(sourceOrder.id, 'Рахунок на оплату')) {
      setNotice('Акт не можна створити без рахунку. Спочатку сформуйте рахунок.');
      return;
    }
    if (kind === 'Видаткова накладна' && sourceSale && sourceSale.items.length === 0) {
      setNotice('Видаткова накладна продажу створюється тільки якщо є товарні позиції.');
      return;
    }
    const gross = sourceOrder ? orderTotals(sourceOrder).total : sourceSale ? saleTotals(sourceSale).total : 0;
    const totals = invoiceTotals(gross, companySettings.vatEnabled);
    const latestAppliedPayment = sourceOrder?.payments.find((payment) => paymentCountsAsApplied(payment));
    const documentStatus: PrintDocument['status'] = kind === 'Акт надання послуг'
      ? 'Роздруковано'
      : kind === 'Видаткова накладна'
        ? 'Роздруковано'
        : kind === 'Рахунок на оплату'
          ? (sourceOrder && orderTotals(sourceOrder).debt > 0 ? 'Очікує оплати' : 'Готово')
          : 'PDF збережено';
    const snapshotPayload = sourceOrder
      ? {
          ...buildServiceOrderDocumentSnapshot(
            kind as 'Рахунок на оплату' | 'Акт надання послуг' | 'Наряд на ремонт' | 'Гарантійний талон',
            sourceOrder,
            products,
            number,
            companySettings,
            clientTaxId(sourceOrder.client, sourceOrder.phone),
            {
              invoiceNumber: findServiceOrderDocument(sourceOrder.id, 'Рахунок на оплату')?.number,
              invoiceDate: findServiceOrderDocument(sourceOrder.id, 'Рахунок на оплату')?.createdAt,
              invoiceId: findServiceOrderDocument(sourceOrder.id, 'Рахунок на оплату')?.id,
              taxInvoiceId: findLatestOrderTaxInvoice(taxInvoices, sourceOrder.id)?.id,
              paymentId: latestAppliedPayment?.id,
            },
          ),
          paymentRef: latestAppliedPayment?.id ?? '',
          paymentMethod: latestAppliedPayment?.method ?? '',
        }
      : undefined;
    const documentId = uid('DOC');
    const document: PrintDocument = {
      id: documentId,
      kind,
      number,
      entityType,
      entityId,
      clientOrSupplier,
      createdAt: today,
      createdBy: activeUser.name,
      version: 1,
      status: documentStatus,
      pdfPath: `/documents/${entityType}/${entityId}/${number}.pdf`,
      amountNet: totals.totalWithoutVAT,
      amountVat: totals.vat,
      amountGross: totals.totalWithVAT,
      paymentId: latestAppliedPayment?.id,
      orderId: sourceOrder?.id,
      invoiceId: kind === 'Рахунок на оплату' ? documentId : sourceOrder ? findServiceOrderDocument(sourceOrder.id, 'Рахунок на оплату')?.id : undefined,
      actId: kind === 'Акт надання послуг' ? documentId : sourceOrder ? findServiceOrderDocument(sourceOrder.id, 'Акт надання послуг')?.id : undefined,
      taxInvoiceId: sourceOrder ? findLatestOrderTaxInvoice(taxInvoices, sourceOrder.id)?.id : undefined,
      waybillId: kind === 'Видаткова накладна' ? documentId : sourceOrder ? findServiceOrderDocument(sourceOrder.id, 'Видаткова накладна')?.id : undefined,
      snapshotData: snapshotPayload ? JSON.stringify(snapshotPayload) : undefined,
      printedAt: documentStatus === 'Роздруковано' ? today : undefined,
      printedBy: documentStatus === 'Роздруковано' ? activeUser.name : undefined,
    };

    setDocuments((current) => [document, ...current]);
    setNotice(`Створено PDF-документ "${kind}" ${number}. Він прив'язаний до ${entityId} і записаний у журнал.`);
    logAction('Створення PDF', entityId, `${kind} ${number}: створено, підготовлено до друку${latestAppliedPayment ? `, платіж ${latestAppliedPayment.id}` : ''}.`);
    window.print();
  }

  function canAcceptShiftBasedPayment(method: PaymentMethod | 'наличные' | 'карта' | 'перевод', silent = false) {
    const normalizedMethod = method === 'наличные' ? 'Готівка' : method === 'карта' ? 'Картка' : method === 'перевод' ? 'Безготівка' : method;
    if (normalizedMethod === 'Безготівка') return true;
    if (cashShift.status !== 'Відкрита') {
      if (!silent) setNotice('Спочатку відкрийте касову зміну.');
      return false;
    }
    return true;
  }

  function registerCashPayment(amount: number, method: PaymentMethod, isRefund = false) {
    setCashShift((current) => {
      const sign = isRefund ? -1 : 1;
      return {
        ...current,
        cashIncome: method === 'Готівка' ? current.cashIncome + amount * sign : current.cashIncome,
        cardIncome: method === 'Картка' ? current.cardIncome + amount * sign : current.cardIncome,
        bankIncome: method === 'Безготівка' ? current.bankIncome + amount * sign : current.bankIncome,
        cashExpense: isRefund && method === 'Готівка' ? current.cashExpense + amount : current.cashExpense,
      };
    });
  }

  function openCashShift(openingCashInput: string) {
    const openingCash = Number(openingCashInput);
    if (!Number.isFinite(openingCash) || openingCash < 0) {
      setNotice('Вкажіть коректний початковий залишок готівки.');
      return false;
    }
    if (cashShift.status === 'Відкрита') {
      setNotice('Касова зміна вже відкрита.');
      return false;
    }
    const nextShift: CashShift = {
      id: `SHIFT-${Date.now()}`,
      openedAt: today,
      openedBy: activeUser.name,
      status: 'Відкрита',
      openingCash,
      cashIncome: 0,
      cardIncome: 0,
      bankIncome: 0,
      cashExpense: 0,
    };
    setCashShift(nextShift);
    logAction('Відкриття зміни', nextShift.id, `Початковий залишок ${money(openingCash)}. Відкрив ${activeUser.name}.`);
    setNotice(`Касову зміну відкрито. Стартова готівка: ${money(openingCash)}.`);
    return true;
  }

  function closeCashShift(actualCashInput: string, comment: string) {
    if (cashShift.status !== 'Відкрита') {
      setNotice('Немає відкритої касової зміни.');
      return false;
    }
    const actualCash = Number(actualCashInput);
    if (!Number.isFinite(actualCash) || actualCash < 0) {
      setNotice('Вкажіть фактичний залишок готівки.');
      return false;
    }
    const expectedCash = expectedCashAmount(cashShift);
    const difference = actualCash - expectedCash;
    if (Math.abs(difference) > 0.009 && !comment.trim()) {
      setNotice('При розбіжності касову зміну можна закрити тільки з коментарем.');
      return false;
    }
    setCashShift((current) => ({
      ...current,
      status: 'Закрита',
      closedAt: today,
      closedBy: activeUser.name,
      actualCash,
      difference,
      closeComment: comment.trim(),
    }));
    logAction('Закриття зміни', cashShift.id, `Очікувано ${money(expectedCash)} · фактично ${money(actualCash)} · різниця ${money(difference)}${comment.trim() ? ` · ${comment.trim()}` : ''}.`);
    if (Math.abs(difference) > 0.009) {
      logAction('Розбіжність по касі', cashShift.id, `${money(difference)} · ${comment.trim() || 'Без коментаря'}.`);
    }
    setNotice(`Зміну ${cashShift.id} закрито. Очікувано: ${money(expectedCash)} · фактично: ${money(actualCash)}.`);
    return true;
  }

  function writeOffInstalledPartsForOrder(order: ServiceOrder, comment: string) {
    let writtenOff = 0;
    order.parts.forEach((part) => {
      if (part.status === 'Встановлено') {
        patchProduct(part.productId, (item) => ({ ...item, installed: Math.max(item.installed - part.qty, 0) }));
        patchOrderPart(order.id, part.id, (item) => ({ ...item, status: 'Списано' }));
        addMovement({
          type: 'Списання',
          productId: part.productId,
          qty: part.qty,
          orderId: order.id,
          batchRefs: part.batchAllocations?.map((item) => `${item.batchId} x${item.qty}`).join(', '),
          unitPrice: part.cost,
          totalAmount: part.cost * part.qty,
          comment,
        });
        logAction('Списання запчастини після оплати', order.id, `${productName(part.productId)}: ${part.qty} шт.`);
        writtenOff += 1;
      }
    });
    return writtenOff;
  }

  function patchProduct(productId: string, updater: (product: Product) => Product) {
    setProducts((current) => current.map((product) => (product.id === productId ? updater(product) : product)));
  }

  function receiveStockIntake(input: StockIntakeInput) {
    const normalizedName = buildNormalizedProductName(input);
    const name = normalizedName || input.name.trim();
    const supplier = input.supplier?.trim() || 'Прямий прихід';
    const qty = Number(input.qty);
    const purchasePrice = Number(input.purchasePrice);
    if (!name || Number.isNaN(qty) || qty <= 0 || Number.isNaN(purchasePrice) || purchasePrice <= 0) {
      setNotice('Для приходу потрібні назва, кількість і закупівельна ціна більше нуля.');
      return;
    }

    const purchaseId = uid('PO');
    const existingProductId = input.appendBarcodeToProductId || input.productId;
    const batchId = uid('BAT');
    let createdProductId: string | undefined;
    const salePrice = Number(input.salePrice) > 0 ? Number(input.salePrice) : Math.round(purchasePrice * 1.35);
    const minStock = Number(input.minStock) > 0 ? Number(input.minStock) : 2;
    const normalizedBarcode = normalizeBarcodeValue(input.barcode?.trim() || '');
    const brand = normalizeWarehouseText(input.brand ?? '');
    const model = normalizeWarehouseText(input.model ?? '');
    const unit = normalizeWarehouseText(input.unit ?? '') || 'шт';
    const storageLocation = normalizeWarehouseText(input.storageLocation ?? '');
    const purchaseDocumentNo = input.purchaseDocumentNo?.trim() || purchaseId;
    const normalizedSku = normalizeWarehouseText(input.sku?.trim() || '');

    if (!existingProductId && !normalizedSku) {
      setNotice('Нова картка товару не може бути створена без SKU.');
      return;
    }
    if (!existingProductId && !name) {
      setNotice('Для нової картки потрібна нормалізована назва товару.');
      return;
    }
    if (!existingProductId && !ensureUniqueSku(products, normalizedSku)) {
      setNotice(`SKU ${normalizedSku} уже існує. Використайте наявну картку або додайте новий штрих-код до неї.`);
      return;
    }
    if (!existingProductId && !input.forceCreateNew) {
      const duplicates = findSimilarProducts(products, { name, sku: normalizedSku, barcode: normalizedBarcode, brand, model });
      if (duplicates.length > 0) {
        setNotice('Можливо, такий товар уже є. Спочатку виберіть наявну картку або додайте штрих-код до неї.');
        return;
      }
    }

    if (existingProductId) {
      patchProduct(existingProductId, (product) => ({
        ...product,
        name: buildNormalizedProductName({ category: input.category ?? product.category, brand: brand || product.brand, model: model || product.model, name: product.name }),
        barcode: product.barcode || normalizedBarcode || product.sku,
        extraBarcodes: normalizedBarcode && normalizedBarcode !== product.barcode
          ? normalizeBarcodeList(...product.extraBarcodes, normalizedBarcode).filter((item) => item !== product.barcode)
          : product.extraBarcodes,
        category: input.category?.trim() || product.category,
        brand: brand || product.brand,
        model: model || product.model,
        unit,
        stock: product.stock + qty,
        cost: purchasePrice,
        price: product.price > 0 ? product.price : salePrice,
        min: Math.max(product.min, minStock),
        storageLocation: storageLocation || product.storageLocation,
        batches: [
          ...product.batches,
          {
            id: batchId,
            productId: existingProductId,
            qtyTotal: qty,
            qtyAvailable: qty,
            purchasePrice,
            purchaseDate: today,
            source: supplier,
            supplier,
            documentNo: purchaseDocumentNo,
            comment: input.appendBarcodeToProductId ? `Додано штрих-код ${normalizedBarcode || 'без коду'} і нову партію.` : 'Прямий прихід на склад.',
          },
        ],
      }));
    } else {
      createdProductId = uid('p');
      const sku = normalizedSku;
      const barcode = normalizedBarcode || uid('BC').replace(/[^A-Z0-9]/gi, '').slice(0, 14);
      const category = input.category?.trim() || 'Складська позиція';
      const newProduct: Product = {
        id: createdProductId,
        name,
        sku,
        barcode,
        extraBarcodes: [],
        category,
        brand,
        model,
        unit,
        cost: purchasePrice,
        price: salePrice,
        stock: qty,
        reserved: 0,
        withEngineer: 0,
        installed: 0,
        min: minStock,
        storageLocation,
        batches: [
          {
            id: batchId,
            productId: createdProductId,
            qtyTotal: qty,
            qtyAvailable: qty,
            purchasePrice,
            purchaseDate: today,
            source: supplier,
            supplier,
            documentNo: purchaseDocumentNo,
            comment: 'Нова позиція створена автоматично при приході.',
          },
        ],
      };
      setProducts((current) => [
        newProduct,
        ...current,
      ]);
    }
    const productId = existingProductId ?? createdProductId;
    if (!productId) return;

    setPurchases((current) => [
      {
        id: purchaseId,
        supplier,
        items: [{ productId, qty, price: purchasePrice, received: qty }],
        status: 'Прибуло',
        orderedAt: today,
        expectedAt: today,
      },
      ...current,
    ]);
    setReceipts((current) => [
      { id: uid('GR'), date: today, supplier, productId, qty, price: purchasePrice, purchaseId },
      ...current,
    ]);
    addMovement({
      type: 'Прихід',
      productId,
      qty,
      purchaseId,
      basis: purchaseDocumentNo,
      batchRefs: batchId,
      unitPrice: purchasePrice,
      totalAmount: qty * purchasePrice,
      comment: input.productId || input.appendBarcodeToProductId
        ? `Прихід нової партії${normalizedBarcode ? ` · штрих-код ${normalizedBarcode}` : ''}.`
        : 'Прийом товару на склад з автоматичним створенням нової номенклатури.',
    });
    logAction('Прийом товару', productId, `${name} · ${qty} ${unit} · ${money(purchasePrice)} · ${supplier}. Партія ${batchId}. Документ ${purchaseDocumentNo}.`);
    setNotice(input.appendBarcodeToProductId ? `Штрих-код додано до ${name}, партію оприбутковано.` : input.productId ? `Оприбутковано ${qty} ${unit} для ${name} новою партією.` : `Створено нову позицію ${name} і оприбутковано ${qty} ${unit}.`);
  }

  function patchOrderPart(orderId: string, partId: string, updater: (part: OrderPart) => OrderPart) {
    setOrders((current) =>
      current.map((order) =>
        order.id === orderId
          ? { ...order, parts: order.parts.map((part) => (part.id === partId ? updater(part) : part)) }
          : order,
      ),
    );
  }

  function patchRequirement(requirementId: string, updater: (requirement: PartRequirement) => PartRequirement) {
    setRequirements((current) => current.map((requirement) => (requirement.id === requirementId ? updater(requirement) : requirement)));
  }

  function patchPurchase(purchaseId: string, updater: (purchase: PurchaseOrder) => PurchaseOrder) {
    setPurchases((current) => current.map((purchase) => (purchase.id === purchaseId ? updater(purchase) : purchase)));
  }

  function updateOrderStatus(orderId: string, status: OrderStatus, comment = 'Статус змінено системою.') {
    setOrders((current) =>
      current.map((order) =>
        order.id === orderId
          ? {
              ...order,
              status,
              statusChangedAt: today,
              statusHistory: [
                { id: uid('H'), oldStatus: order.status, newStatus: status, changedBy: activeUser.name, changedAt: today, comment },
                ...order.statusHistory,
              ],
            }
          : order,
      ),
    );
  }

  function addPartToRepair() {
    const product = products.find((item) => item.id === selectedProductId);
    if (!product || !selectedOrder) return;
    if (selectedOrder.pendingExtraApproval) {
      setNotice('Не можна змінювати суму замовлення, поки додаткові роботи не погоджені клієнтом.');
      return;
    }
    if (selectedOrder.payments.some(paymentCountsAsApplied)) {
      setNotice('Фінансова частина вже зафіксована оплатою. Змінити склад замовлення можна тільки через окреме коригування.');
      return;
    }
    const requestedQty = Math.max(1, qty);
    const orderPartId = uid('OP');

    if (available(product) >= requestedQty) {
      setOrders((current) =>
        current.map((order) =>
          order.id === selectedOrder.id
            ? {
                ...order,
                status: 'В ремонті',
                parts: [...order.parts, { id: orderPartId, productId: product.id, qty: requestedQty, status: 'Зарезервовано', cost: product.cost, price: product.price }],
              }
            : order,
        ),
      );
      patchProduct(product.id, (item) => ({ ...item, reserved: item.reserved + requestedQty }));
      addMovement({ type: 'Резерв', productId: product.id, qty: requestedQty, orderId: selectedOrder.id, comment: 'Резерв створено автоматично при додаванні в ремонт.' });
      logAction('Резерв запчастини', selectedOrder.id, `${product.name}: ${requestedQty} шт.`);
      setNotice(`Запчастину "${product.name}" зарезервовано під ${selectedOrder.id}.`);
      return;
    }

    const requirementId = uid('REQ');
    setRequirements((current) => [{
      id: requirementId,
      orderId: selectedOrder.id,
      productId: product.id,
      qty: requestedQty,
      status: 'Потрібно',
      reason: 'Під замовлення',
      priority: 'Високий',
      requester: activeUser.name,
      comment: `Автоматична потреба під ремонт ${selectedOrder.id}.`,
    }, ...current]);
    setOrders((current) =>
      current.map((order) =>
        order.id === selectedOrder.id
          ? {
              ...order,
              status: 'Очікує запчастину',
              parts: [...order.parts, { id: orderPartId, productId: product.id, qty: requestedQty, status: 'Потрібно', cost: product.cost, price: product.price, requirementId }],
            }
          : order,
      ),
    );
    enqueueClientNotification(selectedOrder, 'Очікування запчастини');
    setNotice(`На складі немає "${product.name}". Створено потребу для ${selectedOrder.id}.`);
    logAction('Створення потреби', selectedOrder.id, `${product.name}: ${requestedQty} шт.`);
  }

  function orderPart(order: ServiceOrder, part: OrderPart) {
    if (order.payments.some(paymentCountsAsApplied)) {
      setNotice('Після оплати не можна змінювати товарну частину замовлення без окремого коригування.');
      return;
    }
    const purchaseId = uid('PO');
    const requirementId = part.requirementId ?? uid('REQ');
    if (!part.requirementId) {
      setRequirements((current) => [{
        id: requirementId,
        orderId: order.id,
        productId: part.productId,
        qty: part.qty,
        status: 'До закупівлі',
        reason: 'Під замовлення',
        priority: 'Високий',
        requester: activeUser.name,
        comment: `Потреба створена менеджером для ${order.id}.`,
      }, ...current]);
    }
    setRequirements((current) => current.map((item) => (item.id === requirementId ? { ...item, status: 'Замовлено' } : item)));
    setPurchases((current) => [
      {
        id: purchaseId,
        supplier: 'Не вибрано',
        items: [{ productId: part.productId, qty: part.qty, price: part.cost, received: 0, requirementId, orderId: order.id }],
        status: 'Нова',
        orderedAt: today,
        expectedAt: '22.04.2026',
        reason: 'Під замовлення',
        priority: 'Високий',
        requestedBy: activeUser.name,
      },
      ...current,
    ]);
    patchOrderPart(order.id, part.id, (item) => ({ ...item, status: 'Замовлено', requirementId, purchaseId }));
    updateOrderStatus(order.id, 'Очікує запчастину');
    enqueueClientNotification(order, 'Очікування запчастини');
    logAction('Створення закупівлі', purchaseId, `Закупівля прив'язана до ремонту ${order.id}.`);
    setNotice(`Створено закупівлю ${purchaseId} і прив'язано її до ${order.id}.`);
  }

  function movePurchase(purchaseId: string) {
    setPurchases((current) => current.map((purchase) => (
      purchase.id === purchaseId
        ? {
            ...purchase,
            status: purchase.status === 'Нова' || purchase.status === 'В роботі' ? 'Замовлено' : 'В дорозі',
            purchasedAt: purchase.purchasedAt ?? today,
          }
        : purchase
    )));
    const purchase = purchases.find((item) => item.id === purchaseId);
    purchase?.items.forEach((item) => {
      if (item.requirementId) {
        setRequirements((current) => current.map((req) => (req.id === item.requirementId ? { ...req, status: 'В дорозі' } : req)));
      }
      if (item.orderId) {
        setOrders((current) =>
          current.map((order) =>
            order.id === item.orderId
              ? { ...order, parts: order.parts.map((part) => (part.purchaseId === purchaseId ? { ...part, status: 'В дорозі' } : part)) }
              : order,
          ),
        );
      }
    });
    logAction('Зміна статусу закупівлі', purchaseId, 'Закупівля переведена в замовлено / в дорозі.');
    setNotice(`Закупівля ${purchaseId} оновлена.`);
  }

  function createManualPurchaseRequest(input: { productId: string; qty: number; reason: PurchaseReason; priority: PurchasePriority; comment: string }) {
    const product = products.find((item) => item.id === input.productId);
    if (!product || input.qty <= 0) {
      setNotice('Не можна створити заявку без товару і кількості.');
      return;
    }
    const requirementId = uid('REQ');
    setRequirements((current) => [
      {
        id: requirementId,
        orderId: '',
        productId: input.productId,
        qty: input.qty,
        status: 'До закупівлі',
        reason: input.reason,
        priority: input.priority,
        requester: activeUser.name,
        comment: input.comment.trim(),
      },
      ...current,
    ]);
    setNotice(`Заявку на закупівлю для ${product.name} створено.`);
  }

  function createPurchaseFromRequirement(requirementId: string) {
    const requirement = requirements.find((item) => item.id === requirementId);
    const product = requirement ? products.find((item) => item.id === requirement.productId) : null;
    if (!requirement || !product) {
      setNotice('Потребу для закупівлі не знайдено.');
      return;
    }
    if (purchases.some((purchase) => purchase.items.some((item) => item.requirementId === requirementId))) {
      setNotice('Для цієї потреби закупівля вже створена.');
      return;
    }
    const purchaseId = uid('PO');
    setPurchases((current) => [
      {
        id: purchaseId,
        supplier: requirement.supplier?.trim() || 'Не вибрано',
        items: [{ productId: requirement.productId, qty: requirement.qty, price: requirement.purchasePrice ?? product.cost, received: 0, requirementId, orderId: requirement.orderId || undefined }],
        status: 'Нова',
        orderedAt: today,
        expectedAt: today,
        reason: requirement.reason ?? 'Вручну',
        priority: requirement.priority ?? 'Середній',
        sourceLink: requirement.sourceLink,
        comment: requirement.comment,
        requestedBy: requirement.requester ?? activeUser.name,
      },
      ...current,
    ]);
    patchRequirement(requirementId, (item) => ({ ...item, status: 'Замовлено' }));
    setNotice(`Закупівлю ${purchaseId} створено.`);
  }

  function updatePurchaseProcurementMeta(purchaseId: string, patch: Partial<PurchaseOrder>) {
    patchPurchase(purchaseId, (purchase) => ({ ...purchase, ...patch }));
    const linkedRequirementIds = purchases.find((purchase) => purchase.id === purchaseId)?.items.map((item) => item.requirementId).filter(Boolean) as string[] | undefined;
    if (linkedRequirementIds?.length) {
      setRequirements((current) => current.map((requirement) => (
        linkedRequirementIds.includes(requirement.id)
          ? {
              ...requirement,
              supplier: typeof patch.supplier === 'string' ? patch.supplier : requirement.supplier,
              purchasePrice: typeof patch.items?.[0]?.price === 'number' ? patch.items[0].price : requirement.purchasePrice,
              sourceLink: typeof patch.sourceLink === 'string' ? patch.sourceLink : requirement.sourceLink,
              comment: typeof patch.comment === 'string' ? patch.comment : requirement.comment,
            }
          : requirement
      )));
    }
  }

  function receivePurchase(purchaseId: string) {
    const purchase = purchases.find((item) => item.id === purchaseId);
    if (!purchase) return;
    purchase.items.forEach((item) => {
      const qtyToReceive = item.qty - item.received;
      if (qtyToReceive <= 0) return;
      patchProduct(item.productId, (product) => ({
        ...product,
        stock: product.stock + qtyToReceive,
        cost: item.price,
        batches: [
          ...product.batches,
          {
            id: uid('BAT'),
            productId: item.productId,
            qtyTotal: qtyToReceive,
            qtyAvailable: qtyToReceive,
            purchasePrice: item.price,
            purchaseDate: today,
            source: purchase.supplier,
            supplier: purchase.supplier,
            documentNo: purchaseId,
            comment: `Партія створена з закупки ${purchaseId}.`,
          },
        ],
      }));
      setReceipts((current) => [
        { id: uid('GR'), date: today, supplier: purchase.supplier, productId: item.productId, qty: qtyToReceive, price: item.price, purchaseId },
        ...current,
      ]);
      addMovement({
        type: 'Прихід',
        productId: item.productId,
        qty: qtyToReceive,
        purchaseId,
        basis: purchaseId,
        batchRefs: `${purchaseId}:${qtyToReceive}`,
        unitPrice: item.price,
        totalAmount: qtyToReceive * item.price,
        comment: 'Прихід за закупівлею. CRM пропонує резерв під замовлення.',
      });
      if (item.requirementId) {
        setRequirements((current) => current.map((req) => (req.id === item.requirementId ? { ...req, status: 'Прибуло' } : req)));
      }
      if (item.orderId) {
        setOrders((current) =>
          current.map((order) =>
            order.id === item.orderId
              ? { ...order, parts: order.parts.map((part) => (part.purchaseId === purchaseId ? { ...part, status: 'Прибуло' } : part)) }
              : order,
          ),
        );
      }
    });
    setPurchases((current) =>
      current.map((item) =>
        item.id === purchaseId
          ? { ...item, status: 'Отримано', receivedAt: today, items: item.items.map((line) => ({ ...line, received: line.qty })) }
          : item,
      ),
    );
    logAction('Отримання закупівлі', purchaseId, 'Закупівлю отримано, створено прихід і FIFO-партії.');
    logAction('Прихід товару', purchaseId, 'Оформлено прихід за закупівлею.');
    setNotice(`Прихід за ${purchaseId} оформлено. Тепер можна зарезервувати запчастину під ремонт.`);
  }

  function reserveArrived(order: ServiceOrder, part: OrderPart) {
    const product = products.find((item) => item.id === part.productId);
    if (!product || available(product) < part.qty) {
      setNotice('Резерв неможливий: доступного залишку недостатньо.');
      return;
    }
    patchProduct(part.productId, (item) => ({ ...item, reserved: item.reserved + part.qty }));
    patchOrderPart(order.id, part.id, (item) => ({ ...item, status: 'Зарезервовано' }));
    addMovement({ type: 'Резерв', productId: part.productId, qty: part.qty, orderId: order.id, comment: 'Резерв після приходу товару.' });
    logAction('Резерв після приходу', order.id, `${product.name}: ${part.qty} шт.`);
    setNotice(`Запчастину "${product.name}" зарезервовано під ${order.id}.`);
  }

  function issueToEngineer(order: ServiceOrder, part: OrderPart) {
    const product = products.find((item) => item.id === part.productId);
    if (!product || part.status !== 'Зарезервовано' || product.reserved < part.qty || product.stock < part.qty) {
      setNotice('Видача неможлива: потрібен активний резерв.');
      return;
    }
    const fifo = fifoConsumeBatches(product, part.qty);
    if (!fifo) {
      setNotice('Видача неможлива: по партіях не вистачає залишку, перевірте склад.');
      return;
    }
    patchProduct(part.productId, (item) => ({
      ...item,
      stock: item.stock - part.qty,
      reserved: item.reserved - part.qty,
      withEngineer: item.withEngineer + part.qty,
      batches: fifo.batches,
    }));
    patchOrderPart(order.id, part.id, (item) => ({
      ...item,
      status: 'Видано інженеру',
      cost: Math.round(fifo.totalCost / part.qty),
      batchAllocations: fifo.allocations,
    }));
    addMovement({
      type: 'Видача інженеру',
      productId: part.productId,
      qty: part.qty,
      orderId: order.id,
      batchRefs: fifo.allocations.map((line) => `${line.batchId} x${line.qty}`).join(', '),
      unitPrice: part.qty > 0 ? Math.round(fifo.totalCost / part.qty) : 0,
      totalAmount: fifo.totalCost,
      comment: `Видано інженеру ${order.engineer}. FIFO: ${fifo.allocations.map((line) => `${line.batchId} x${line.qty}`).join(', ')}.`,
    });
    logAction('Видача інженеру', order.id, `${product.name}: ${part.qty} шт. видано ${order.engineer}. FIFO ${fifo.allocations.map((line) => `${line.batchId}/${line.qty}`).join(', ')}.`);
    setNotice(`Запчастину видано інженеру ${order.engineer}. Списання по партіях виконано FIFO.`);
  }

  function markInstalled(order: ServiceOrder, part: OrderPart) {
    const product = products.find((item) => item.id === part.productId);
    if (!product || part.status !== 'Видано інженеру' || product.withEngineer < part.qty) {
      setNotice('Встановлення неможливе: запчастина ще не видана інженеру.');
      return;
    }
    patchProduct(part.productId, (item) => ({ ...item, withEngineer: item.withEngineer - part.qty, installed: item.installed + part.qty }));
    patchOrderPart(order.id, part.id, (item) => ({ ...item, status: 'Встановлено' }));
    addMovement({ type: 'Встановлення', productId: part.productId, qty: part.qty, orderId: order.id, comment: 'Інженер відмітив запчастину як встановлену.' });
    logAction('Встановлення запчастини', order.id, `${product.name}: ${part.qty} шт.`);
    setNotice('Запчастину встановлено. При закритті ремонту CRM спише собівартість документом.');
  }

  function closeOrder(order: ServiceOrder) {
    if (order.pendingExtraApproval) {
      setNotice('Закрити ремонт не можна: клієнт ще не погодив додаткові роботи.');
      return;
    }
    const notInstalled = order.parts.some((part) => !['Встановлено', 'Списано'].includes(part.status));
    if (notInstalled) {
      setNotice('Закрити ремонт не можна: є запчастини не в кінцевому статусі.');
      return;
    }
    if (orderTotals(order).debt > 0) {
      setNotice('Закрити і видати ремонт не можна: є борг клієнта.');
      return;
    }
    const actIsSigned = documents.some((document) => document.entityType === 'service_order' && document.entityId === order.id && document.kind === 'Акт надання послуг' && document.status === 'Підписано');
    if (!actIsSigned) {
      setNotice('Закрити замовлення не можна: акт виконаних робіт ще не підписано.');
      return;
    }
    order.parts.forEach((part) => {
      if (part.status === 'Встановлено') {
        patchProduct(part.productId, (item) => ({ ...item, installed: Math.max(item.installed - part.qty, 0) }));
        patchOrderPart(order.id, part.id, (item) => ({ ...item, status: 'Списано' }));
        addMovement({
          type: 'Списання',
          productId: part.productId,
          qty: part.qty,
          orderId: order.id,
          batchRefs: part.batchAllocations?.map((item) => `${item.batchId} x${item.qty}`).join(', '),
          unitPrice: part.cost,
          totalAmount: part.cost * part.qty,
          comment: 'Автоматичне списання при закритті ремонту.',
        });
        logAction('Списання при закритті ремонту', order.id, `${productName(part.productId)}: ${part.qty} шт.`);
      }
    });
    updateOrderStatus(order.id, 'Закрито');
    const warranty = addDocumentIfMissing('Гарантійний талон', 'service_order', order.id, order.client, 'Створено', order);
    enqueueClientNotification(order, 'Видача');
    logAction('Закриття і видача ремонту', order.id, 'Ремонт закрито, пристрій видано клієнту.');
    if (warranty) logAction('Гарантію оформлено', order.id, `Створено ${warranty.number} після закриття ремонту.`);
    setNotice(`Ремонт ${order.id} закрито. Встановлені запчастини списано документом.`);
  }

  function issueReadyOrder(order: ServiceOrder) {
    if (order.pendingExtraApproval) {
      setNotice('Видати замовлення не можна: є несогласовані додаткові роботи.');
      return;
    }
    const actStatus = getOrderActStatus(documents, order.id);
    const releaseBlockReason = getBlockReason(order, actStatus);
    if (!canRelease(order, actStatus) || releaseBlockReason) {
      const reasonMessage = releaseBlockReason === 'Є борг'
        ? 'Видати замовлення не можна: є борг клієнта. Спочатку прийміть оплату або окремо зафіксуйте дозвіл керівника.'
        : releaseBlockReason === 'Платіж не підтверджено'
          ? `Видати замовлення не можна: ${pendingPaymentReason(order) || 'платіж не підтверджено'}.`
          : releaseBlockReason === 'Акт не підписано'
            ? 'Видати замовлення не можна: акт не підписано.'
            : 'Видати замовлення можна тільки після статусу "Готовий до видачі".';
      setNotice(reasonMessage);
      return;
    }
    const hasIssueDocument = documents.some((document) => (
      document.entityType === 'service_order'
      && document.entityId === order.id
      && ['Акт надання послуг', 'Акт виконаних робіт', 'Акт видачі', 'Акт технічного стану'].includes(document.kind)
    ));
    if (!hasIssueDocument) {
      setNotice('Видати замовлення не можна: спочатку має бути сформований акт.');
      return;
    }
    const actIsSigned = documents.some((document) => document.entityType === 'service_order' && document.entityId === order.id && document.kind === 'Акт надання послуг' && document.status === 'Підписано');
    if (!actIsSigned) {
      setNotice('Видати замовлення не можна: акт виконаних робіт ще не підписано.');
      return;
    }
    const unfinishedParts = order.parts.some((part) => !['Встановлено', 'Списано', 'Повернення'].includes(part.status));
    if (unfinishedParts) {
      setNotice('Видати замовлення не можна: є незавершені запчастини або резерви.');
      return;
    }
    if (!order.locationCode || order.locationStatus === 'У інженера') {
      setNotice('Видати замовлення не можна: пристрій має бути фізично повернений у закріплену ячейку.');
      return;
    }
    order.parts.forEach((part) => {
      if (part.status === 'Встановлено') {
        patchProduct(part.productId, (item) => ({ ...item, installed: Math.max(item.installed - part.qty, 0) }));
        patchOrderPart(order.id, part.id, (item) => ({ ...item, status: 'Списано' }));
        addMovement({
          type: 'Списання',
          productId: part.productId,
          qty: part.qty,
          orderId: order.id,
          batchRefs: part.batchAllocations?.map((item) => `${item.batchId} x${item.qty}`).join(', '),
          unitPrice: part.cost,
          totalAmount: part.cost * part.qty,
          comment: 'Автоматичне списання при видачі і закритті замовлення менеджером.',
        });
        logAction('Списання при видачі ремонту', order.id, `${productName(part.productId)}: ${part.qty} шт.`);
      }
    });
    setOrders((current) => current.map((item) => {
      if (item.id !== order.id) return item;
      return {
        ...item,
        status: 'Закрито',
        statusChangedAt: today,
        actIssuedAt: item.legalEntity ? (item.actIssuedAt ?? today) : item.actIssuedAt,
        statusHistory: [
          {
            id: uid('H'),
            oldStatus: item.status,
            newStatus: 'Закрито',
            changedBy: activeUser.name,
            changedAt: today,
            comment: `Менеджер видав замовлення клієнту. Дата видачі та користувач зафіксовані: ${activeUser.name}.`,
          },
          ...item.statusHistory,
        ],
      };
    }));
    const warranty = addDocumentIfMissing('Гарантійний талон', 'service_order', order.id, order.client, 'Створено', order);
    enqueueClientNotification(order, 'Видача');
    logAction('Видача і закриття замовлення', order.id, `Видав ${activeUser.name}. Дата видачі: ${today}.`);
    if (warranty) logAction('Гарантію оформлено', order.id, `Створено ${warranty.number} після видачі клієнту.`);
    moveOrder(order.id, 'OUT');
    setNotice(`${order.id}: замовлення видано і закрито. Дата видачі та користувач зафіксовані.`);
    setSuggestedDocumentAction({
      orderId: order.id,
      kind: 'Гарантійний талон',
      title: 'Видача завершена',
      text: 'Після видачі система одразу відкриває гарантійний талон.',
      autoOpen: true,
    });
  }

  function validateOrderTransition(order: ServiceOrder, nextStatus: OrderStatus) {
    if (nextStatus === 'Не підлягає ремонту') {
      if (!roleCanSetOrderStatus(activeUser.role, nextStatus)) {
        return `Роль "${activeUser.role}" не має права встановлювати статус "${nextStatus}".`;
      }
      if (['Закрито', 'Видано', 'Скасовано'].includes(order.status)) {
        return 'Не можна змінити вже виданий, закритий або скасований заказ на "Не підлягає ремонту".';
      }
      return '';
    }
    const allowedNext = orderStatusFlow[order.status] ?? [];
    if (order.status !== nextStatus && !allowedNext.includes(nextStatus) && activeUser.role !== 'Руководитель' && activeUser.role !== 'Адміністратор') {
      return `Перехід "${order.status}" -> "${nextStatus}" не дозволений бізнес-процесом.`;
    }
    if (!roleCanSetOrderStatus(activeUser.role, nextStatus)) {
      return `Роль "${activeUser.role}" не має права встановлювати статус "${nextStatus}".`;
    }
    if (order.pendingExtraApproval && ['В ремонті', 'На тестуванні', 'Готовий до видачі', 'Видано', 'Закрито'].includes(nextStatus)) {
      return 'Не можна продовжити ремонт або закрити замовлення, поки додаткові роботи не погоджені клієнтом.';
    }
    if (nextStatus === 'Погоджено' && !order.diagnosisResult) {
      return 'Не можна погодити ремонт без результату діагностики.';
    }
    if (nextStatus === 'Очікує запчастину' && order.parts.length === 0) {
      return 'Не можна очікувати запчастину без створеної потреби або доданої деталі.';
    }
    if (nextStatus === 'В ремонті') {
      if (!order.engineer) return 'Не можна передати в ремонт без призначеного інженера.';
      if (!order.engineerAcceptedAt) return 'Інженер має спочатку натиснути "Взяти в роботу".';
    }
    if (nextStatus === 'Готовий до видачі') {
      const unfinishedParts = order.parts.some((part) => !['Встановлено', 'Списано', 'Повернення'].includes(part.status));
      if (unfinishedParts) return 'Не можна позначити готовим: є незавершені запчастини.';
      if (!order.engineerWorkCompletedAt) return 'Не можна перевести в "Готово": інженер ще не завершив роботи.';
      if (!order.locationCode) return 'Не можна перевести в "Готово": у замовлення має бути закріплена ячейка.';
      if (order.locationStatus === 'У інженера') return 'Не можна перевести в "Готово": інженер має спочатку повернути пристрій у закріплену ячейку.';
    }
    if (nextStatus === 'Видано' && order.status !== 'Готовий до видачі') {
      return 'Видати можна тільки після статусу "Готовий до видачі".';
    }
    if (nextStatus === 'Закрито' && order.status !== 'Видано') {
      return 'Закрити заказ можна тільки після фізичної видачі пристрою.';
    }
    return '';
  }

  function changeOrderStatus(order: ServiceOrder, nextStatus: OrderStatus, comment: string) {
    const error = validateOrderTransition(order, nextStatus);
    if (error) {
      setNotice(error);
      return;
    }
    if (nextStatus === 'Не підлягає ремонту' && comment.trim().length < 10) {
      setNotice('Потрібно вказати причину: чому пристрій не підлягає ремонту.');
      return;
    }
    if (nextStatus === 'Скасовано') {
      order.parts.forEach((part) => {
        if (part.status === 'Зарезервовано') {
          patchProduct(part.productId, (product) => ({ ...product, reserved: Math.max(product.reserved - part.qty, 0) }));
          addMovement({ type: 'Зняття резерву', productId: part.productId, qty: part.qty, orderId: order.id, comment: 'Ремонт скасовано, резерв знято.' });
        }
      });
    }
    if (nextStatus === 'Готовий до видачі') {
      if (order.works.length === 0) {
        setNotice('Не можна завершити ремонт: не додано жодної виконаної роботи.');
        return;
      }
      if (orderTotals(order).works <= 0) {
        setNotice('Не можна завершити ремонт: не вказано суму виконаних робіт.');
        return;
      }
      updateOrderStatus(order.id, nextStatus, comment);
      enqueueClientNotification({ ...order, status: nextStatus }, 'Готово до видачі');
      const createdKinds = createSmartReadyDocuments(order);
      logAction('Зміна статусу ремонту', order.id, `${order.status} -> ${nextStatus}. ${comment}`);
      setNotice(`Інженер натиснув "Готово". ${order.id}: акт, рахунок${order.parts.length > 0 ? ' і видаткова' : ''} перевірені автоматично. Створено: ${createdKinds.length ? createdKinds.join(', ') : 'документи вже були'}.`);
      setSuggestedDocumentAction({
        orderId: order.id,
        kind: 'Акт надання послуг',
        title: 'Замовлення готове',
        text: 'Після переходу у готовність система одразу відкриває акт виконаних робіт.',
        autoOpen: true,
      });
      return;
    }
    if (nextStatus === 'Не підлягає ремонту') {
      updateOrderStatus(order.id, 'Не підлягає ремонту', comment);
      enqueueClientNotification(order, 'Готово до видачі');
      logAction('Не підлягає ремонту', order.id, `Причина: ${comment}. Замовлення виведено з активних ремонтів і готове до видачі клієнту.`);
      setNotice(`${order.id}: ремонт неможливий. Замовлення виведено з активних ремонтів і готове до видачі.`);
      return;
    }
    if (nextStatus === 'Видано') {
      updateOrderStatus(order.id, 'Видано', comment);
      enqueueClientNotification(order, 'Видача');
      logAction('Видача пристрою', order.id, 'Пристрій фізично передано клієнту.');
      return;
    }
    if (nextStatus === 'Закрито') {
      closeOrder(order);
      return;
    }
    updateOrderStatus(order.id, nextStatus, comment);
    const event = notificationEventForStatus(nextStatus);
    if (event) enqueueClientNotification({ ...order, status: nextStatus }, event);
    logAction('Зміна статусу ремонту', order.id, `${order.status} -> ${nextStatus}. ${comment}`);
    setNotice(`${order.id}: статус змінено на "${nextStatus}".`);
  }

  function acceptOrderWork(order: ServiceOrder) {
    if (activeUser.role !== 'Інженер' || order.engineer !== activeUser.name) {
      setNotice('Взяти в роботу може тільки призначений інженер.');
      return;
    }
    if (!order.locationCode) {
      setNotice('Взяти в ремонт неможливо: у замовлення немає закріпленої ячейки.');
      return;
    }
    setOrders((current) => current.map((item) => {
      if (item.id !== order.id) return item;
      const nextStatus: OrderStatus = item.status === 'Прийнято' ? 'В ремонті' : item.status;
      return {
        ...item,
        engineerAcceptedAt: today,
        takenInRepairAt: today,
        takenInRepairBy: activeUser.name,
        takenFromLocation: item.locationCode,
        locationStatus: 'У інженера',
        status: nextStatus,
        statusChangedAt: nextStatus !== item.status ? today : item.statusChangedAt,
        statusHistory: nextStatus !== item.status
          ? [{ id: uid('H'), oldStatus: item.status, newStatus: nextStatus, changedBy: activeUser.name, changedAt: today, comment: `Інженер натиснув "Взяти в ремонт". Пристрій взято з ячейки ${item.locationCode}.` }, ...item.statusHistory]
          : item.statusHistory,
      };
    }));
    setOrderUnits((current) => current.map((unit) => (
      unit.orderId !== order.id ? unit : { ...unit, locationCode: order.locationCode, status: 'В ремонті', lastActionAt: today }
    )));
    logAction('Інженер взяв заказ', order.id, `${activeUser.name} підтвердив роботу. Взято з ячейки ${order.locationCode}.`);
    setNotice(`${order.id}: інженер взяв замовлення в ремонт із ячейки ${order.locationCode}. Ячейка залишається закріпленою за цим замовленням.`);
  }

  function markEngineerWorkCompleted(order: ServiceOrder) {
    setOrders((current) => current.map((item) => (
      item.id !== order.id
        ? item
        : {
            ...item,
            engineerWorkCompletedAt: today,
            engineerWorkCompletedBy: activeUser.name,
            statusHistory: [
              { id: uid('H'), oldStatus: item.status, newStatus: item.status, changedBy: activeUser.name, changedAt: today, comment: 'Роботу виконано інженером.' },
              ...item.statusHistory,
            ],
          }
    )));
    logAction('Роботу виконано інженером', order.id, `${activeUser.name} зафіксував завершення робіт.`);
    setNotice(`${order.id}: інженер зафіксував, що роботу виконано.`);
  }

  function returnOrderToCellReady(order: ServiceOrder) {
    if (!window.confirm(`Підтверджую, що фізично поклав замовлення ${order.id} в ячейку ${order.locationCode}?`)) {
      return;
    }
    if (!order.engineerAcceptedAt) {
      setNotice('Не можна повернути в ячейку як готове: замовлення ще не було взято в ремонт.');
      return;
    }
    if (!order.locationCode) {
      setNotice('Повернення неможливе: у замовлення немає закріпленої ячейки.');
      return;
    }
    if (order.locationStatus !== 'У інженера') {
      setNotice('Замовлення вже знаходиться у своїй ячейці.');
      return;
    }
    if (order.works.length === 0) {
      setNotice('Не можна завершити ремонт: не додано жодної роботи.');
      return;
    }
    if (orderTotals(order).works <= 0) {
      setNotice('Не можна завершити ремонт: не вказана сума робіт.');
      return;
    }
    setOrders((current) => current.map((item) => (
      item.id !== order.id
        ? item
        : {
            ...item,
            locationStatus: 'У комірці',
            returnedToCellAt: today,
            returnedToCellBy: activeUser.name,
          }
    )));
    setOrderUnits((current) => current.map((unit) => (
      unit.orderId !== order.id ? unit : { ...unit, locationCode: order.locationCode, status: 'На полці', lastActionAt: today }
    )));
    markEngineerWorkCompleted(order);
    changeOrderStatus(
      {
        ...order,
        locationStatus: 'У комірці',
        engineerWorkCompletedAt: order.engineerWorkCompletedAt ?? today,
        engineerWorkCompletedBy: order.engineerWorkCompletedBy ?? activeUser.name,
      },
      'Готовий до видачі',
      `Інженер завершив ремонт і повернув замовлення в ячейку ${order.locationCode}.`,
    );
    logAction('Повернення в ячейку', order.id, `${activeUser.name} повернув пристрій у ячейку ${order.locationCode}.`);
    logAction('Сигнал менеджеру', order.id, `Замовлення готове і повернене в ячейку ${order.locationCode}.`);
    setNotice(`Заказ ${order.id} готов и возвращён в ячейку ${order.locationCode}. Клиенту автоматически сформировано уведомление о готовности к выдаче.`);
  }

  function cancelOrder(order: ServiceOrder, reason?: string, comment?: string) {
    const reasonText = reason?.trim();
    const commentText = comment?.trim();
    if (!reasonText) {
      if (!window.confirm(`Скасувати замовлення ${order.id}? Історія не буде видалена.`)) return;
    }
    const auditDetail = [reasonText ? `Причина: ${reasonText}` : 'Причина не вказана', commentText].filter(Boolean).join(' · ');
    createOrderVersion(order, `Скасування замовлення${reasonText ? `: ${reasonText}` : ''}`);
    updateOrderStatus(order.id, 'Скасовано', `Замовлення скасовано окремою дією.${auditDetail ? ` ${auditDetail}.` : ''}`);
    appendSimpleOrderActivity(order.id, 'Замовлення скасовано', auditDetail);
    enqueueClientNotification(order, 'Нагадування', 'SMS', true, `Замовлення ${order.id} скасовано. Якщо це помилка, зверніться в сервісний центр.`);
    logAction('Скасування замовлення', order.id, auditDetail || 'Замовлення скасовано без видалення попередньої історії.');
    setNotice(`${order.id}: замовлення скасовано. Історія та документи збережені.`);
  }

  function refundOrder(order: ServiceOrder, amountInput?: number, reason?: string, comment?: string) {
    const paid = confirmedOrderPaymentsTotal(order);
    if (paid <= 0) {
      setNotice('Повернення коштів неможливе: оплат по замовленню немає.');
      return;
    }
    const refundAmount = Math.max(0, Math.min(Number(amountInput ?? paid), paid));
    if (!Number.isFinite(refundAmount) || refundAmount <= 0) {
      setNotice('Вкажіть коректну суму повернення.');
      return;
    }
    const reasonText = reason?.trim();
    const commentText = comment?.trim();
    if (!reasonText) {
      if (!window.confirm(`Повернути клієнту ${money(refundAmount)} по замовленню ${order.id}?`)) return;
    }
    const lastConfirmedPayment = order.payments.find((payment) => paymentCountsAsApplied(payment) && payment.amount > 0);
    const refundMethod = lastConfirmedPayment?.method ?? 'Готівка';
    const refundRecord: Payment = {
      id: uid('PAY'),
      date: today,
      amount: -refundAmount,
      method: refundMethod,
      type: 'Повернення коштів',
      transactionNo: uid('REF'),
      acceptedBy: activeUser.name,
      status: 'Підтверджено',
      orderId: order.id,
      comment: [reasonText || 'Повернення коштів', commentText].filter(Boolean).join(' · '),
    };
    const simpleRefundRecord: SimpleOrderPaymentRecord = {
      id: refundRecord.id,
      client: order.client,
      entityType: 'order',
      entityId: order.id,
      clientTaxId: clientTaxId(order.client, order.phone),
      orderId: order.id,
      amount: -refundAmount,
      method: refundMethod === 'Картка' ? 'карта' : refundMethod === 'Безготівка' ? 'перевод' : 'наличные',
      paymentKind: 'оплата',
      status: 'подтвержден',
      date: today,
      direction: 'outgoing',
      purpose: [reasonText || 'Повернення коштів', commentText].filter(Boolean).join(' · '),
    };
    createOrderVersion(order, `Повернення коштів${reasonText ? `: ${reasonText}` : ''}`);
    persistOrdersUpdate((current) => current.map((item) => (
      item.id === order.id
        ? {
            ...item,
            payments: [refundRecord, ...item.payments],
            activityLog: [
              { id: uid('ACT'), date: today, action: 'Повернення коштів', detail: `${money(refundAmount)}${reasonText ? ` · ${reasonText}` : ''}${commentText ? ` · ${commentText}` : ''}` },
              ...(item.activityLog ?? []),
            ],
          }
        : item
    )));
    setSimplePayments((current) => {
      const updated = [simpleRefundRecord, ...current];
      saveSimplePaymentsToStorage(updated);
      return updated;
    });
    logAction('Повернення коштів', order.id, `${money(refundAmount)}${reasonText ? ` · ${reasonText}` : ''}${commentText ? ` · ${commentText}` : ''}`);
    setNotice(`${order.id}: повернення коштів оформлено. Це окрема подія і не стирає попередні оплати.`);
  }

  function cancelOrderAct(order: ServiceOrder) {
    const act = documents.find((document) => document.entityType === 'service_order' && document.entityId === order.id && document.kind === 'Акт надання послуг');
    if (!act) {
      setNotice('Скасувати акт неможливо: акт ще не створено.');
      return;
    }
    if (!window.confirm(`Скасувати акт ${act.number} для ${order.id}? Попередня версія залишиться в історії.`)) return;
    createOrderVersion(order, 'Скасування акта та підготовка нової версії');
    setDocuments((current) => current.map((document) => (
      document.id === act.id
        ? { ...document, status: 'Скасовано' }
        : document
    )));
    logAction('Скасування акта', order.id, `Акт ${act.number} скасовано. Попередня версія збережена.`);
    setNotice(`${order.id}: акт скасовано. Для нових змін буде сформовано нову версію.`);
  }

  function reopenOrder(order: ServiceOrder, reason?: string, comment?: string) {
    const reasonText = reason?.trim();
    const commentText = comment?.trim();
    const canReopen = ['Готовий до видачі', 'Очікує оплати'].includes(order.status) || Boolean(order.engineerWorkCompletedAt || order.returnedToCellAt);
    if (!canReopen) {
      setNotice('Повернути в роботу можна тільки з готового стану або зі стану очікування видачі.');
      return;
    }
    if (!reasonText) {
      if (!window.confirm(`Повторно відкрити замовлення ${order.id}? Це створить нову версію.`)) return;
    }
    createOrderVersion(order, `Повернення в роботу${reasonText ? `: ${reasonText}` : ''}`);
    persistOrdersUpdate((current) => current.map((item) => (
      item.id === order.id
        ? {
            ...item,
            status: 'В ремонті',
            statusChangedAt: today,
            engineerWorkCompletedAt: undefined,
            engineerWorkCompletedBy: undefined,
            returnedToCellAt: undefined,
            returnedToCellBy: undefined,
            locationStatus: 'У інженера',
            statusHistory: [
              { id: uid('H'), oldStatus: item.status, newStatus: 'В ремонті', changedBy: activeUser.name, changedAt: today, comment: [reasonText ? `Причина: ${reasonText}` : 'Повернення в роботу', commentText].filter(Boolean).join(' · ') },
              ...(item.statusHistory ?? []),
            ],
            activityLog: [
              { id: uid('ACT'), date: today, action: 'Повернення в роботу', detail: [reasonText, commentText].filter(Boolean).join(' · ') || 'Без додаткового коментаря' },
              ...(item.activityLog ?? []),
            ],
          }
        : item
    )));
    logAction('Повернення в роботу', order.id, [reasonText, commentText].filter(Boolean).join(' · ') || 'Замовлення повернуто в роботу.');
    setNotice(`${order.id}: замовлення повернуто в роботу. Попередні документи та історія збережені.`);
  }

  function reassignEngineer(order: ServiceOrder) {
    const engineers = users.filter((user) => user.role === 'Інженер');
    const currentIndex = engineers.findIndex((user) => user.name === order.engineer);
    const nextEngineer = engineers[(currentIndex + 1) % engineers.length] ?? engineers[0];
    if (!nextEngineer) return;
    setOrders((current) => current.map((item) => (item.id === order.id ? { ...item, assignedTo: nextEngineer.id, engineer: nextEngineer.name, engineerAcceptedAt: undefined } : item)));
    prependActionLog({ id: uid('LOG'), date: today, user: activeUser.name, role: activeUser.role, action: 'Зміна інженера', entity: order.id, comment: `${order.engineer} -> ${nextEngineer.name}` });
    setNotice(`${order.id}: інженера змінено на ${nextEngineer.name}.`);
  }

  function recordManagerCall(order: ServiceOrder) {
    logAction('Дзвінок клієнту', order.id, `${activeUser.name} зв’язався з клієнтом ${order.client} (${order.phone}).`);
    setNotice(`${order.id}: зафіксовано дзвінок клієнту о ${today}.`);
  }

  function addOrderPayment(order: ServiceOrder, amount: number, method: PaymentMethod, comment: string) {
    const totals = orderTotals(order);
    if (amount <= 0) {
      setNotice('Сума оплати має бути більшою за 0.');
      return;
    }
    if (totals.debt <= 0) {
      setNotice('Фінансова частина замовлення вже закрита: боргу немає.');
      return;
    }
    if (!canAcceptShiftBasedPayment(method)) return;
    const amountToApply = Math.min(amount, totals.debt);
    const status: Payment['status'] = method === 'Безготівка' ? 'Очікує надходження' : method === 'Картка' ? 'Проведено' : 'Підтверджено';
    const type = paymentTypeFor(amountToApply, totals.debt || amountToApply, comment === 'Передплата' ? 'Передплата' : 'Доплата');
    const transactionNo = method === 'Картка' ? uid('TERM') : method === 'Безготівка' ? uid('BANK') : uid('CASH');
    const paidInFull = totals.debt - amountToApply <= 0;
    const canAutoReadyAfterPayment = paidInFull && method === 'Готівка' && ['Готовий до видачі', 'Очікує оплати', 'Очікує клієнта'].includes(order.status);
    setOrders((current) =>
      current.map((item) =>
        item.id === order.id
          ? {
              ...item,
              status: canAutoReadyAfterPayment ? 'Готовий до видачі' : item.status,
              statusChangedAt: canAutoReadyAfterPayment ? today : item.statusChangedAt,
              vatStatus: item.vatStatus ?? 'Очікує ПН',
              debtSince: item.debtSince ?? today,
              payments: [...item.payments, { id: uid('PAY'), date: today, amount: amountToApply, method, type, transactionNo, acceptedBy: activeUser.name, orderId: order.id, comment, status, cashShiftId: method === 'Безготівка' ? undefined : cashShift.id, countedAtShiftId: method === 'Безготівка' ? undefined : cashShift.id }],
              statusHistory: canAutoReadyAfterPayment && item.status !== 'Готовий до видачі'
                ? [{ id: uid('H'), oldStatus: item.status, newStatus: 'Готовий до видачі', changedBy: 'Система', changedAt: today, comment: 'Оплата 100%. CRM автоматично підтвердила готовність до видачі.' }, ...item.statusHistory]
                : item.statusHistory,
            }
          : item,
      ),
    );
    if (method !== 'Безготівка') {
      registerCashPayment(amountToApply, method);
      if (paidInFull) {
        const writtenOff = writeOffInstalledPartsForOrder(order, 'Автоматичне списання встановлених запчастин після повної оплати ремонту.');
        if (writtenOff > 0) logAction('Автосписання після оплати', order.id, `Списано позицій: ${writtenOff}.`);
      }
    }
    logAction('Прийом оплати ремонту', order.id, `${method}: ${money(amountToApply)}. ${comment}. Статус: ${status}.`);
    logAction('НДС подія', order.id, 'Оплата зафіксована: дані готові до формування НН за правилами системи.');
    if (canAutoReadyAfterPayment) logAction('Автостатус після оплати', order.id, 'Оплата 100%: CRM перевела заказ у "Готовий до видачі".');
    setNotice(method === 'Безготівка'
      ? `${comment}: безготівкову оплату ${money(amountToApply)} зафіксовано для ${order.id}. Потрібне підтвердження бухгалтера.`
      : `${comment}: прийнято ${money(amountToApply)} для ${order.id}. Залишок ${money(Math.max(totals.debt - amountToApply, 0))}.`);
  }

  function oneClickManagerIssue(order: ServiceOrder, payment?: { method: PaymentMethod; amount?: number; partial?: boolean }) {
    if (order.pendingExtraApproval) {
      setNotice('Видати замовлення не можна: є несогласовані додаткові роботи.');
      return;
    }
    if (!order.locationCode || order.locationStatus === 'У інженера') {
      setNotice('Видати замовлення не можна: пристрій має бути фізично повернений у закріплену ячейку.');
      return;
    }
    const hasUnfinishedParts = order.parts.some((part) => !['Встановлено', 'Списано', 'Повернення'].includes(part.status));
    if (hasUnfinishedParts) {
      setNotice('Видати замовлення не можна: є незавершені запчастини або резерви.');
      return;
    }
    const totals = orderTotals(order);
    if (payment && totals.works <= 0) {
      setNotice('Прийняти оплату не можна: не вказана сума робіт.');
      return;
    }

    ensureOrderDocumentRecord('Акт надання послуг', order);
    const hasAct = documents.some((document) => document.entityType === 'service_order' && document.entityId === order.id && document.kind === 'Акт надання послуг');
    if (!hasAct) {
      setNotice(`Для ${order.id} створено акт. CRM підготувала його автоматично, повторіть видачу ще раз.`);
      setSuggestedDocumentAction({
        orderId: order.id,
        kind: 'Акт надання послуг',
        title: 'Акт підготовлено',
        text: 'CRM автоматично створила акт для видачі.',
        autoOpen: true,
      });
      return;
    }
    const actSigned = documents.some((document) => document.entityType === 'service_order' && document.entityId === order.id && document.kind === 'Акт надання послуг' && document.status === 'Підписано');
    if (!actSigned) signOrderAct(order);

    const amountToApply = payment ? Math.min(payment.amount && payment.amount > 0 ? payment.amount : totals.debt, totals.debt) : 0;
    const willRemainDebt = Math.max(totals.debt - amountToApply, 0);
    const isPartial = Boolean(payment?.partial) || willRemainDebt > 0;

    if (payment && amountToApply > 0) {
      addOrderPayment(order, amountToApply, payment.method, payment.partial ? 'Часткова оплата при видачі' : 'Оплата при видачі');
      if (payment.method !== 'Безготівка') {
        printDocument('Касовий чек', 'service_order', order.id, order.client);
      }
    }

    if (payment && (payment.partial || willRemainDebt > 0)) {
      setSuggestedDocumentAction(null);
      return;
    }

    const issuedOrder: ServiceOrder = {
      ...order,
      payments: payment && amountToApply > 0
        ? [...order.payments, { id: uid('PAY'), date: today, amount: amountToApply, method: payment.method, type: 'Доплата', transactionNo: payment.method === 'Картка' ? uid('TERM') : payment.method === 'Безготівка' ? uid('BANK') : uid('CASH'), acceptedBy: activeUser.name, orderId: order.id, comment: 'Оплата при видачі', status: payment.method === 'Безготівка' ? 'Очікує надходження' : payment.method === 'Картка' ? 'Проведено' : 'Підтверджено' }]
        : order.payments,
      status: 'Готовий до видачі',
    };
    issuedOrder.parts.forEach((part) => {
      if (part.status === 'Встановлено') {
        patchProduct(part.productId, (item) => ({ ...item, installed: Math.max(item.installed - part.qty, 0) }));
        patchOrderPart(issuedOrder.id, part.id, (item) => ({ ...item, status: 'Списано' }));
        addMovement({
          type: 'Списання',
          productId: part.productId,
          qty: part.qty,
          orderId: issuedOrder.id,
          batchRefs: part.batchAllocations?.map((item) => `${item.batchId} x${item.qty}`).join(', '),
          unitPrice: part.cost,
          totalAmount: part.cost * part.qty,
          comment: 'Автоматичне списання при видачі через міні-касу.',
        });
        logAction('Списання при видачі ремонту', issuedOrder.id, `${productName(part.productId)}: ${part.qty} шт.`);
      }
    });
    setOrders((current) => current.map((item) => {
      if (item.id !== issuedOrder.id) return item;
      return {
        ...item,
        status: 'Закрито',
        statusChangedAt: today,
        locationStatus: 'Видано',
        actIssuedAt: item.legalEntity ? (item.actIssuedAt ?? today) : item.actIssuedAt,
        statusHistory: [
          {
            id: uid('H'),
            oldStatus: item.status,
            newStatus: 'Закрито',
            changedBy: activeUser.name,
            changedAt: today,
            comment: `Менеджер видав замовлення клієнту через міні-касу. Дата видачі: ${today}.`,
          },
          ...item.statusHistory,
        ],
      };
    }));
    const warranty = addDocumentIfMissing('Гарантійний талон', 'service_order', issuedOrder.id, issuedOrder.client, 'Створено', issuedOrder);
    enqueueClientNotification(issuedOrder, 'Видача');
    logAction('Видача і закриття замовлення', issuedOrder.id, `Видав ${activeUser.name}. Дата видачі: ${today}.`);
    if (warranty) logAction('Гарантію оформлено', issuedOrder.id, `Створено ${warranty.number} після видачі клієнту.`);
    moveOrder(issuedOrder.id, 'OUT');
    setSuggestedDocumentAction({
      orderId: issuedOrder.id,
      kind: 'Гарантійний талон',
      title: 'Видача завершена',
      text: 'Після видачі система одразу відкриває гарантійний талон.',
      autoOpen: true,
    });
    setNotice(`${issuedOrder.id}: замовлення видано і закрито. Каса, гарантія та SMS виконані автоматично.`);
    const nextOrder = orders.find((item) => item.id !== order.id && !['Закрито', 'Видано', 'Скасовано'].includes(item.status));
    if (nextOrder) setSelectedOrderId(nextOrder.id);
  }

  function transferOrderToBas(order: ServiceOrder) {
    const orderDocuments = documents.filter((document) => document.entityType === 'service_order' && document.entityId === order.id);
    const hasAct = orderDocuments.some((document) => ['Акт надання послуг', 'Акт виконаних робіт', 'Акт видачі', 'Акт технічного стану'].includes(document.kind));
    const hasInvoice = orderDocuments.some((document) => document.kind === 'Рахунок на оплату');
    const hasDelivery = order.parts.length === 0 || orderDocuments.some((document) => document.kind === 'Видаткова накладна');
    const hasPayment = order.payments.some(paymentCountsAsApplied);
    if (!hasAct || !hasInvoice || !hasDelivery || !hasPayment) {
      setNotice('Передача в BAS недоступна: потрібні акт, рахунок, видаткова для товарів і підтверджена оплата.');
      return;
    }
    logAction('Передача в BAS', order.id, `Передано: акт, рахунок, ${order.parts.length ? 'видаткова, ' : ''}оплата. Документів: ${orderDocuments.length}.`);
    setNotice(`${order.id}: дані підготовлено і передано в BAS: акт, рахунок, ${order.parts.length ? 'видаткова, ' : ''}оплата.`);
  }

  function confirmBankPayment(paymentId: string) {
    let confirmed: Payment | undefined;
    let sourceOrder: ServiceOrder | undefined;
    let remainingAfterConfirmation = 0;
    let alreadyCountedAtShiftId = '';
    setOrders((current) => current.map((order) => {
      const payment = order.payments.find((item) => item.id === paymentId && paymentNeedsConfirmation(item));
      if (!payment) return order;
      confirmed = payment;
      sourceOrder = order;
      alreadyCountedAtShiftId = payment.countedAtShiftId ?? '';
      const currentRemaining = clientOrderDebt(order);
      const nextRemaining = Math.max(currentRemaining - payment.amount, 0);
      remainingAfterConfirmation = nextRemaining;
      return {
        ...order,
        status: nextOrderStatusAfterAutoPayment(order, nextRemaining),
        statusChangedAt: nextOrderStatusAfterAutoPayment(order, nextRemaining) !== order.status ? today : order.statusChangedAt,
        payments: order.payments.map((item) => (item.id === paymentId ? { ...item, status: 'Підтверджено', confirmedBy: activeUser.name, confirmedAt: today, countedAtShiftId: item.countedAtShiftId ?? (item.method === 'Безготівка' ? 'BANK_CONFIRMED' : cashShift.id) } : item)),
        activityLog: [
          { id: uid('ACT'), date: today, action: 'Платіж підтверджено', detail: `${money(payment.amount)} · залишок ${money(nextRemaining)}` },
          ...(order.activityLog ?? []),
        ],
      };
    }));
    setSimplePayments((current) => current.map((payment) => (
      payment.id === paymentId
        ? { ...payment, status: 'подтвержден', countedAtShiftId: payment.countedAtShiftId ?? (payment.method === 'перевод' ? 'BANK_CONFIRMED' : cashShift.id) }
        : payment
    )));
    setSales((current) => current.map((sale) => {
      const payment = sale.payments.find((item) => item.id === paymentId && paymentNeedsConfirmation(item));
      if (payment && !confirmed) confirmed = payment;
      if (payment && !alreadyCountedAtShiftId) alreadyCountedAtShiftId = payment.countedAtShiftId ?? '';
      return {
        ...sale,
        payments: sale.payments.map((item) => (item.id === paymentId && paymentNeedsConfirmation(item) ? { ...item, status: 'Підтверджено', confirmedBy: activeUser.name, confirmedAt: today, countedAtShiftId: item.countedAtShiftId ?? (item.method === 'Безготівка' ? 'BANK_CONFIRMED' : cashShift.id) } : item)),
      };
    }));
    if (!confirmed) {
      setNotice('Платіж не знайдено або він уже підтверджений.');
      return;
    }
    if (!alreadyCountedAtShiftId) {
      registerCashPayment(confirmed.amount, confirmed.method);
    }
    if (sourceOrder && orderTotals(sourceOrder).debt - confirmed.amount <= 0) {
      writeOffInstalledPartsForOrder(sourceOrder, 'Автоматичне списання встановлених запчастин після підтвердження безготівкової оплати.');
    }
    logAction(
      confirmed.method === 'Картка' ? 'Підтвердження термінала' : confirmed.method === 'Безготівка' ? 'Підтвердження безготівки' : 'Підтвердження платежу',
      confirmed.orderId ?? confirmed.saleId ?? paymentId,
      `${money(confirmed.amount)} підтвердив ${activeUser.name}.`,
    );
    if (sourceOrder?.id && remainingAfterConfirmation <= 0) {
      appendSimpleOrderActivity(sourceOrder.id, 'Борг закрито', `Платіж ${money(confirmed.amount)} підтверджено`);
      logAction('Закриття боргу', sourceOrder.id, `Борг закрито платежем ${money(confirmed.amount)}.`);
      if (nextOrderStatusAfterAutoPayment(sourceOrder, remainingAfterConfirmation) === 'Готовий до видачі') {
        enqueueClientNotification({ ...sourceOrder, status: 'Готовий до видачі' }, 'Готово до видачі');
      }
    }
    setNotice(`Платіж ${money(confirmed.amount)} підтверджено.`);
  }

  function cancelPayment(paymentId: string, reason: string, comment = '') {
    const reasonText = reason.trim() || 'Скасування платежу';
    let cancelledPayment: Payment | undefined;
    let affectedEntity = paymentId;
    let registeredAtShift = false;
    let registeredMethod: PaymentMethod | null = null;
    let updatedOrdersSnapshot = orders;
    let updatedSalesSnapshot = sales;

    updatedOrdersSnapshot = orders.map((order) => {
      const source = order.payments.find((payment) => payment.id === paymentId && payment.amount > 0 && paymentCountsAsApplied(payment));
      if (!source) return order;
      cancelledPayment = source;
      affectedEntity = order.id;
      registeredAtShift = Boolean(source.countedAtShiftId);
      registeredMethod = source.method;
      const reversal: Payment = {
        id: uid('PAY'),
        date: today,
        amount: -Math.abs(source.amount),
        method: source.method,
        type: 'Повернення коштів',
        transactionNo: uid('VOID'),
        acceptedBy: activeUser.name,
        status: 'Скасовано',
        orderId: order.id,
        comment: [reasonText, comment.trim()].filter(Boolean).join(' · '),
        cashShiftId: source.method === 'Безготівка' ? undefined : cashShift.id,
        countedAtShiftId: source.method === 'Безготівка' ? undefined : cashShift.id,
        cancelledAt: today,
        cancelledBy: activeUser.name,
        refundReason: reasonText,
      };
      const nextRemaining = clientOrderDebt(order) + Math.abs(source.amount);
      return {
        ...order,
        status: ['Видано', 'Закрито', 'Скасовано'].includes(order.status) ? order.status : nextOrderStatusAfterAutoPayment(order, nextRemaining),
        statusChangedAt: ['Видано', 'Закрито', 'Скасовано'].includes(order.status) ? order.statusChangedAt : today,
        payments: [reversal, ...order.payments],
        activityLog: [
          { id: uid('ACT'), date: today, action: 'Платіж скасовано', detail: `${money(source.amount)} · ${reasonText}${comment.trim() ? ` · ${comment.trim()}` : ''}` },
          ...(order.activityLog ?? []),
        ],
      };
    });
    updatedSalesSnapshot = sales.map((sale) => {
      const source = sale.payments.find((payment) => payment.id === paymentId && payment.amount > 0 && paymentCountsAsApplied(payment));
      if (!source || cancelledPayment) return sale;
      cancelledPayment = source;
      affectedEntity = sale.id;
      registeredAtShift = Boolean(source.countedAtShiftId);
      registeredMethod = source.method;
      const reversal: Payment = {
        id: uid('PAY'),
        date: today,
        amount: -Math.abs(source.amount),
        method: source.method,
        type: 'Повернення коштів',
        transactionNo: uid('VOID'),
        acceptedBy: activeUser.name,
        status: 'Скасовано',
        saleId: sale.id,
        comment: [reasonText, comment.trim()].filter(Boolean).join(' · '),
        cashShiftId: source.method === 'Безготівка' ? undefined : cashShift.id,
        countedAtShiftId: source.method === 'Безготівка' ? undefined : cashShift.id,
        cancelledAt: today,
        cancelledBy: activeUser.name,
        refundReason: reasonText,
      };
      return {
        ...sale,
        payments: [reversal, ...sale.payments],
      };
    });

    if (!cancelledPayment || !registeredMethod) {
      setNotice('Для скасування доступні тільки вже проведені платежі.');
      return;
    }
    const safeCancelledPayment = cancelledPayment;
    const cancelledOrderId = safeCancelledPayment.orderId;
    if (registeredMethod !== 'Безготівка' && !canAcceptShiftBasedPayment(registeredMethod)) return;

    setOrders(updatedOrdersSnapshot);
    setSales(updatedSalesSnapshot);
    if (cancelledOrderId) {
      setSimplePayments((current) => [
        {
          id: uid('SMPAY'),
          client: orders.find((order) => order.id === cancelledOrderId)?.client ?? '',
          entityType: 'order',
          entityId: cancelledOrderId,
          clientTaxId: (() => {
            const sourceOrder = orders.find((order) => order.id === cancelledOrderId);
            return sourceOrder ? clientTaxId(sourceOrder.client, sourceOrder.phone) : undefined;
          })(),
          orderId: cancelledOrderId,
          amount: -Math.abs(safeCancelledPayment.amount),
          method: safeCancelledPayment.method === 'Картка' ? 'карта' : safeCancelledPayment.method === 'Безготівка' ? 'перевод' : 'наличные',
          paymentKind: 'оплата',
          status: 'скасовано',
          date: today,
          acceptedBy: activeUser.name,
          cashShiftId: safeCancelledPayment.method === 'Безготівка' ? undefined : cashShift.id,
          countedAtShiftId: safeCancelledPayment.method === 'Безготівка' ? undefined : cashShift.id,
          cancelledAt: today,
          cancelledBy: activeUser.name,
          refundReason: reasonText,
        },
        ...current,
      ]);
    }
    if (registeredAtShift) {
      registerCashPayment(Math.abs(safeCancelledPayment.amount), registeredMethod, true);
    }
    logAction('Скасування платежу', affectedEntity, `${money(Math.abs(safeCancelledPayment.amount))} · ${registeredMethod} · ${reasonText}${comment.trim() ? ` · ${comment.trim()}` : ''}`);
    setNotice(`Платіж скасовано: ${money(Math.abs(safeCancelledPayment.amount))}.`);
  }

  function refundPayment(paymentId: string, amountInput: string, reason: string, method: PaymentMethod) {
    const refundAmount = Number(amountInput);
    const reasonText = reason.trim() || 'Повернення коштів';
    if (!Number.isFinite(refundAmount) || refundAmount <= 0) {
      setNotice('Вкажіть коректну суму повернення.');
      return;
    }
    if (method !== 'Безготівка' && !canAcceptShiftBasedPayment(method)) return;
    let sourcePayment: Payment | undefined;
    let affectedEntity = paymentId;
    const updatedOrders = orders.map((order) => {
      const source = order.payments.find((payment) => payment.id === paymentId && payment.amount > 0 && paymentCountsAsApplied(payment));
      if (!source) return order;
      if (refundAmount > source.amount) {
        setNotice('Сума повернення не може перевищувати суму платежу.');
        return order;
      }
      sourcePayment = source;
      affectedEntity = order.id;
      const reversal: Payment = {
        id: uid('PAY'),
        date: today,
        amount: -refundAmount,
        method,
        type: 'Повернення коштів',
        transactionNo: uid('REF'),
        acceptedBy: activeUser.name,
        status: 'Повернення',
        orderId: order.id,
        comment: reasonText,
        cashShiftId: method === 'Безготівка' ? undefined : cashShift.id,
        countedAtShiftId: method === 'Безготівка' ? undefined : cashShift.id,
        cancelledAt: today,
        cancelledBy: activeUser.name,
        refundReason: reasonText,
      };
      const nextRemaining = clientOrderDebt(order) + refundAmount;
      return {
        ...order,
        status: ['Видано', 'Закрито', 'Скасовано'].includes(order.status) ? order.status : nextOrderStatusAfterAutoPayment(order, nextRemaining),
        statusChangedAt: ['Видано', 'Закрито', 'Скасовано'].includes(order.status) ? order.statusChangedAt : today,
        payments: [reversal, ...order.payments],
        activityLog: [
          { id: uid('ACT'), date: today, action: 'Повернення коштів', detail: `${money(refundAmount)} · ${reasonText}` },
          ...(order.activityLog ?? []),
        ],
      };
    });
    const updatedSales = sales.map((sale) => {
      const source = sale.payments.find((payment) => payment.id === paymentId && payment.amount > 0 && paymentCountsAsApplied(payment));
      if (!source || sourcePayment) return sale;
      if (refundAmount > source.amount) {
        setNotice('Сума повернення не може перевищувати суму платежу.');
        return sale;
      }
      sourcePayment = source;
      affectedEntity = sale.id;
      const reversal: Payment = {
        id: uid('PAY'),
        date: today,
        amount: -refundAmount,
        method,
        type: 'Повернення коштів',
        transactionNo: uid('REF'),
        acceptedBy: activeUser.name,
        status: 'Повернення',
        saleId: sale.id,
        comment: reasonText,
        cashShiftId: method === 'Безготівка' ? undefined : cashShift.id,
        countedAtShiftId: method === 'Безготівка' ? undefined : cashShift.id,
        cancelledAt: today,
        cancelledBy: activeUser.name,
        refundReason: reasonText,
      };
      return {
        ...sale,
        payments: [reversal, ...sale.payments],
      };
    });
    if (!sourcePayment) {
      setNotice('Для повернення доступні тільки вже проведені платежі.');
      return;
    }
    const safeSourcePayment = sourcePayment;
    const sourceOrderId = safeSourcePayment.orderId;
    setOrders(updatedOrders);
    setSales(updatedSales);
    if (sourceOrderId) {
      setSimplePayments((current) => [
        {
          id: uid('SMPAY'),
          client: orders.find((order) => order.id === sourceOrderId)?.client ?? '',
          entityType: 'order',
          entityId: sourceOrderId,
          clientTaxId: (() => {
            const sourceOrder = orders.find((order) => order.id === sourceOrderId);
            return sourceOrder ? clientTaxId(sourceOrder.client, sourceOrder.phone) : undefined;
          })(),
          orderId: sourceOrderId,
          amount: -refundAmount,
          method: method === 'Картка' ? 'карта' : method === 'Безготівка' ? 'перевод' : 'наличные',
          paymentKind: 'оплата',
          status: 'повернення',
          date: today,
          acceptedBy: activeUser.name,
          cashShiftId: method === 'Безготівка' ? undefined : cashShift.id,
          countedAtShiftId: method === 'Безготівка' ? undefined : cashShift.id,
          cancelledAt: today,
          cancelledBy: activeUser.name,
          refundReason: reasonText,
        },
        ...current,
      ]);
    }
    registerCashPayment(refundAmount, method, true);
    logAction('Повернення коштів', affectedEntity, `${money(refundAmount)} · ${method} · ${reasonText}`);
    setNotice(`Повернення оформлено: ${money(refundAmount)}.`);
  }

  function returnServicePart(order: ServiceOrder, part: OrderPart, destination: 'На склад' | 'Брак') {
    if (order.payments.some(paymentCountsAsApplied)) {
      setNotice('Після оплати не можна змінювати склад замовлення без окремого коригування.');
      return;
    }
    if (!['Зарезервовано', 'Видано інженеру', 'Встановлено'].includes(part.status)) {
      setNotice('Повернення неможливе для цього статусу запчастини.');
      return;
    }
    if (part.status === 'Зарезервовано') {
      patchProduct(part.productId, (item) => ({ ...item, reserved: Math.max(item.reserved - part.qty, 0) }));
      addMovement({ type: 'Зняття резерву', productId: part.productId, qty: part.qty, orderId: order.id, comment: 'Запчастина знята з ремонту до видачі інженеру.' });
    }
    if (part.status === 'Видано інженеру') {
      patchProduct(part.productId, (item) => ({
        ...item,
        withEngineer: Math.max(item.withEngineer - part.qty, 0),
        stock: destination === 'На склад' ? item.stock + part.qty : item.stock,
        batches: destination === 'На склад' && part.batchAllocations ? restoreBatchAllocations(item, part.batchAllocations) : item.batches,
      }));
      addMovement({
        type: destination === 'На склад' ? 'Повернення' : 'Списання',
        productId: part.productId,
        qty: part.qty,
        orderId: order.id,
        batchRefs: part.batchAllocations?.map((item) => `${item.batchId} x${item.qty}`).join(', '),
        unitPrice: part.cost,
        totalAmount: part.cost * part.qty,
        comment: destination === 'На склад' ? 'Деталь не підійшла, повернута на склад.' : 'Деталь не підійшла, переведена в брак.',
      });
    }
    if (part.status === 'Встановлено') {
      patchProduct(part.productId, (item) => ({
        ...item,
        installed: Math.max(item.installed - part.qty, 0),
        stock: destination === 'На склад' ? item.stock + part.qty : item.stock,
        batches: destination === 'На склад' && part.batchAllocations ? restoreBatchAllocations(item, part.batchAllocations) : item.batches,
      }));
      addMovement({
        type: destination === 'На склад' ? 'Повернення' : 'Списання',
        productId: part.productId,
        qty: part.qty,
        orderId: order.id,
        batchRefs: part.batchAllocations?.map((item) => `${item.batchId} x${item.qty}`).join(', '),
        unitPrice: part.cost,
        totalAmount: part.cost * part.qty,
        comment: destination === 'На склад' ? 'Встановлена деталь знята і повернута на склад.' : 'Встановлена деталь знята і списана в брак.',
      });
    }
    patchOrderPart(order.id, part.id, (item) => ({ ...item, status: 'Повернення', price: 0, batchAllocations: destination === 'На склад' ? [] : item.batchAllocations }));
    logAction('Повернення в сервісі', order.id, `${productName(part.productId)}: ${part.qty} шт., напрямок: ${destination}.`);
    setNotice('Запчастину знято із сервісного замовлення, собівартість ремонту перерахована.');
  }

  function updateSaleStatus(sale: Sale): SaleStatus {
    const totals = saleTotals(sale);
    const hasReserved = sale.items.some((item) => item.status === 'Зарезервовано');
    const hasIssued = sale.items.length > 0 && sale.items.every((item) => ['Видано', 'Повернення', 'Скасовано'].includes(item.status));
    if (sale.status === 'Скасовано' || sale.status === 'Повернення' || sale.status === 'Закрито') return sale.status;
    if (hasIssued && totals.debt === 0) return 'Закрито';
    if (hasIssued) return 'Видано';
    if (totals.paid > 0 && totals.debt > 0) return 'Частково оплачено';
    if (totals.paid >= totals.total && totals.total > 0) return 'Оплачено';
    if (hasReserved) return totals.paid > 0 ? 'Частково оплачено' : 'Зарезервовано';
    return sale.items.length > 0 ? 'Очікує оплати' : 'Чернетка';
  }

  function patchSale(saleId: string, updater: (sale: Sale) => Sale) {
    setSales((current) => current.map((sale) => (sale.id === saleId ? updater(sale) : sale)));
  }

  function addSaleItem() {
    const product = products.find((item) => item.id === selectedSaleProductId);
    if (!product || !selectedSale) return;
    const line: SaleItem = { id: uid('SI'), productId: product.id, qty: Math.max(1, saleQty), price: product.price, cost: product.cost, discount: 0, status: 'Чернетка' };
    patchSale(selectedSale.id, (sale) => {
      const next = { ...sale, items: [...sale.items, line] };
      return { ...next, status: updateSaleStatus(next) };
    });
    logAction('Додавання товару в продаж', selectedSale.id, `${product.name}: ${line.qty} шт.`);
    setNotice(`Товар "${product.name}" додано в продаж ${selectedSale.id}.`);
  }

  function reserveSale(sale: Sale) {
    let blocked = '';
    sale.items.forEach((item) => {
      if (item.status !== 'Чернетка') return;
      const product = products.find((entry) => entry.id === item.productId);
      if (!product || available(product) < item.qty) {
        blocked = `Недостатньо залишку для "${product?.name ?? item.productId}".`;
        return;
      }
      patchProduct(item.productId, (entry) => ({ ...entry, reserved: entry.reserved + item.qty }));
      addMovement({ type: 'Резерв', productId: item.productId, qty: item.qty, orderId: sale.id, comment: 'Резерв товару під продаж.' });
    });
    if (blocked) {
      setNotice(blocked);
      return;
    }
    patchSale(sale.id, (item) => {
      const next = { ...item, items: item.items.map((line) => (line.status === 'Чернетка' ? { ...line, status: 'Зарезервовано' as const } : line)) };
      return { ...next, status: updateSaleStatus(next) };
    });
    logAction('Резерв продажу', sale.id, 'Товари зарезервовано під продаж.');
    setNotice(`Товари зарезервовано під продаж ${sale.id}.`);
  }

  function acceptSalePayment(sale: Sale, amount: number, method: PaymentMethod, comment: string) {
    if (amount <= 0) return;
    if (!canAcceptShiftBasedPayment(method)) return;
    const type = paymentTypeFor(amount, saleTotals(sale).debt || amount, comment === 'Часткова оплата' ? 'Часткова оплата' : 'Повна оплата');
    const transactionNo = method === 'Картка' ? uid('TERM') : method === 'Безготівка' ? uid('BANK') : uid('CASH');
    const status: Payment['status'] = method === 'Безготівка' ? 'Очікує надходження' : method === 'Картка' ? 'Проведено' : 'Підтверджено';
    patchSale(sale.id, (item) => {
      const next = { ...item, payments: [...item.payments, { id: uid('PAY'), date: today, amount, method, type, transactionNo, acceptedBy: activeUser.name, saleId: sale.id, comment, status, cashShiftId: method === 'Безготівка' ? undefined : cashShift.id, countedAtShiftId: method === 'Безготівка' ? undefined : cashShift.id }] };
      return { ...next, status: updateSaleStatus(next) };
    });
    if (method !== 'Безготівка') registerCashPayment(amount, method);
    logAction('Прийом оплати продажу', sale.id, `${method}: ${money(amount)}. ${comment}`);
    setNotice(method === 'Безготівка' ? `${comment}: безготівку ${money(amount)} зафіксовано для продажу ${sale.id}. Потрібне підтвердження бухгалтера.` : `${comment}: прийнято ${money(amount)} для продажу ${sale.id}.`);
  }

  function issueSale(sale: Sale) {
    const totals = saleTotals(sale);
    if (totals.debt > 0) {
      setNotice('Видати товар не можна: продаж не оплачений повністю.');
      return;
    }
    const hasUnreserved = sale.items.some((item) => item.status !== 'Зарезервовано');
    if (hasUnreserved) {
      setNotice('Видати товар не можна: усі позиції мають бути зарезервовані.');
      return;
    }
    const nextProducts = products.map((product) => ({ ...product, batches: product.batches.map((batch) => ({ ...batch })) }));
    const issuedItems = new Map<string, SaleItem>();
    const saleMovements: Array<{ productId: string; qty: number; cost: number; batchAllocations?: BatchAllocation[]; comment: string }> = [];

    for (const item of sale.items) {
      const productIndex = nextProducts.findIndex((product) => product.id === item.productId);
      if (productIndex === -1) {
        setNotice(`Не вдалося знайти товар ${item.productId} для видачі продажу.`);
        return;
      }
      const product = nextProducts[productIndex];
      const fifo = fifoConsumeBatches(product, item.qty);
      if (!fifo) {
        setNotice(`Не вдалося списати ${product.name}: по партіях недостатньо залишку для видачі.`);
        return;
      }
      nextProducts[productIndex] = {
        ...product,
        stock: product.stock - item.qty,
        reserved: Math.max(product.reserved - item.qty, 0),
        batches: fifo.batches,
      };
      issuedItems.set(item.id, {
        ...item,
        status: 'Видано',
        cost: Math.round(fifo.totalCost / item.qty),
        batchAllocations: fifo.allocations,
      });
      saleMovements.push({
        productId: item.productId,
        qty: item.qty,
        cost: Math.round(fifo.totalCost / item.qty),
        batchAllocations: fifo.allocations,
        comment: `Підтверджена видача клієнту, товар списано зі складу по FIFO. Партії: ${fifo.allocations.map((batch) => batch.batchId).join(', ')}.`,
      });
    }

    setProducts(nextProducts);
    saleMovements.forEach((movement) => addMovement({
      type: 'Продаж',
      productId: movement.productId,
      qty: movement.qty,
      orderId: sale.id,
      batchRefs: movement.batchAllocations?.map((item) => `${item.batchId} x${item.qty}`).join(', '),
      unitPrice: movement.cost,
      totalAmount: movement.cost * movement.qty,
      comment: movement.comment,
    }));
    patchSale(sale.id, (item) => {
      const next = {
        ...item,
        items: item.items.map((line) => issuedItems.get(line.id) ?? line),
      };
      return { ...next, status: updateSaleStatus(next) };
    });
    logAction('Видача продажу', sale.id, 'Товар видано клієнту, склад списано документом по FIFO.');
    setNotice(`Продаж ${sale.id} видано і закрито. Система списала товар по партіях FIFO.`);
  }

  function cancelSale(sale: Sale) {
    sale.items.forEach((item) => {
      if (item.status === 'Зарезервовано') {
        patchProduct(item.productId, (product) => ({ ...product, reserved: Math.max(product.reserved - item.qty, 0) }));
        addMovement({ type: 'Зняття резерву', productId: item.productId, qty: item.qty, orderId: sale.id, comment: 'Продаж скасовано до видачі.' });
      }
    });
    patchSale(sale.id, (item) => ({ ...item, status: 'Скасовано', items: item.items.map((line) => ({ ...line, status: 'Скасовано' as const })) }));
    logAction('Скасування продажу', sale.id, 'Продаж скасовано до видачі, резерв знято.');
    setNotice(`Продаж ${sale.id} скасовано, резерв знято.`);
  }

  function returnSale(sale: Sale) {
    const issued = sale.items.find((item) => item.status === 'Видано');
    if (!issued) {
      setNotice('Повернення можливе тільки після видачі товару.');
      return;
    }
    const refund = Math.max(issued.price * issued.qty - issued.discount, 0);
    setProducts((current) => current.map((product) => {
      if (product.id !== issued.productId) return product;
      return {
        ...product,
        stock: product.stock + issued.qty,
        batches: issued.batchAllocations ? restoreBatchAllocations(product, issued.batchAllocations) : product.batches,
      };
    }));
    addMovement({
      type: 'Повернення',
      productId: issued.productId,
      qty: issued.qty,
      orderId: sale.id,
      batchRefs: issued.batchAllocations?.map((item) => `${item.batchId} x${item.qty}`).join(', '),
      unitPrice: issued.cost,
      totalAmount: issued.cost * issued.qty,
      comment: issued.batchAllocations?.length ? `Повернення товару клієнтом, товар повернено на склад у партії: ${issued.batchAllocations.map((batch) => batch.batchId).join(', ')}.` : 'Повернення товару клієнтом, товар повернено на склад.',
    });
    patchSale(sale.id, (item) => ({
      ...item,
      status: 'Повернення',
      items: item.items.map((line) => (line.id === issued.id ? { ...line, status: 'Повернення' as const } : line)),
      returns: [...item.returns, { id: uid('RET'), saleId: sale.id, productId: issued.productId, qty: issued.qty, reason: 'Повернення клієнтом', refund, destination: 'На склад' }],
      payments: [...item.payments, { id: uid('PAY'), date: today, amount: -refund, method: 'Готівка', type: 'Повернення коштів', transactionNo: uid('REF'), acceptedBy: activeUser.name, saleId: sale.id, comment: 'Повернення коштів клієнту' }],
    }));
    registerCashPayment(refund, 'Готівка', true);
    logAction('Повернення продажу', sale.id, `${productName(issued.productId)}: ${issued.qty} шт., повернення коштів ${money(refund)}.`);
    setNotice(`Оформлено повернення за ${sale.id}. Повернення коштів: ${money(refund)}.`);
  }

  const managerStatusOptions: OrderStatus[] = ['Готовий до видачі', 'Очікує оплати', 'Очікує клієнта', 'Не підлягає ремонту'];
  if (!sessionUserRecord) {
    return (
      <EmployeeLoginScreen
        phone={loginPhone}
        code={loginCode}
        challenge={loginChallenge}
        error={loginError}
        hint={loginHint}
        onPhoneChange={setLoginPhone}
        onCodeChange={setLoginCode}
        onRequestCode={requestPhoneCode}
        onVerifyCode={verifyPhoneCode}
        onBack={resetPhoneLogin}
      />
    );
  }
  return (
    <div className="app-shell">
      <main className="workspace">
        <header className="topbar">
          <button className="icon-button mobile-only" onClick={() => setShowMenu(!showMenu)} aria-label="Відкрити меню">
            {showMenu ? <X size={20} /> : <Menu size={20} />}
          </button>
          <div className="brand workspace-brand">
            <div className="brand-mark">СА</div>
            <div>
              <strong>СПЕКТР-АС CRM</strong>
              <span>{roleDisplay(activeUser.role)}</span>
            </div>
          </div>
          {!hideGlobalSearch && (
            <div className="search-box">
              <Search size={18} />
              <input
                value={query}
                onChange={(event) => {
                  setQuery(event.target.value);
                  setShowGlobalSearchResults(Boolean(event.target.value.trim()));
                }}
                onKeyDown={(event) => {
                  if (event.key === 'Enter') {
                    event.preventDefault();
                    submitGlobalSearch();
                  }
                  if (event.key === 'Escape') {
                    setShowGlobalSearchResults(false);
                  }
                }}
                placeholder="Пошук: замовлення / клієнт / телефон / фірма / пристрій"
              />
              {showGlobalSearchResults && (
                <div className="search-dropdown">
                  {globalSearchResults.map((result) => (
                    <button key={`${result.type}-${result.id}`} type="button" className="search-dropdown-row" onClick={() => openGlobalSearchResult(result)}>
                      <strong>{result.title}</strong>
                      <span>{result.subtitle}</span>
                    </button>
                  ))}
                  {globalSearchResults.length === 0 && <div className="search-dropdown-empty">Нічого не знайдено</div>}
                </div>
              )}
            </div>
          )}
          <div style={{ position: 'relative' }}>
            <button
              className="icon-button notification-button"
              type="button"
              aria-label="Центр повідомлень"
              title="Повідомлення"
              onClick={() => setIsNotificationsOpen((current) => !current)}
            >
              <Bell size={19} />
              {unreadInternalMessages > 0 && <span>{unreadInternalMessages}</span>}
            </button>
            {isNotificationsOpen && (
              <div
                style={{
                  position: 'absolute',
                  top: 'calc(100% + 8px)',
                  right: 0,
                  width: '340px',
                  maxHeight: '360px',
                  overflowY: 'auto',
                  background: '#fff',
                  border: '1px solid #dbe3ef',
                  borderRadius: '16px',
                  boxShadow: '0 18px 40px rgba(15, 23, 42, 0.18)',
                  padding: '12px',
                  zIndex: 30,
                }}
              >
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px' }}>
                  <strong>Повідомлення</strong>
                  <button type="button" className="icon-button" aria-label="Закрити повідомлення" onClick={() => setIsNotificationsOpen(false)}>
                    <X size={16} />
                  </button>
                </div>
                {visibleInternalMessages.length > 0 ? (
                  <div style={{ display: 'grid', gap: '10px' }}>
                    {visibleInternalMessages.slice(0, 8).map((message) => (
                      <div key={message.id} style={{ border: '1px solid #e2e8f0', borderRadius: '12px', padding: '10px 12px', display: 'grid', gap: '6px' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', gap: '12px', alignItems: 'start' }}>
                          <div style={{ display: 'grid', gap: '2px' }}>
                            <strong>{message.createdBy || 'Система'}</strong>
                            <span style={{ fontSize: '12px', color: '#64748b' }}>{getInternalMessageAuthorRole(message, users)}</span>
                          </div>
                          <span style={{ fontSize: '12px', color: '#64748b', textAlign: 'right' }}>{message.createdAt}</span>
                        </div>
                        <div style={{ fontSize: '14px', color: '#0f172a' }}>{message.text}</div>
                        <div style={{ fontSize: '12px', color: '#64748b' }}>
                          {message.orderId ? `Замовлення ${message.orderId} · ` : ''}
                          Строк: {message.dueAt}
                        </div>
                      </div>
                    ))}
                  </div>
                ) : (
                  <div className="empty-state">Новых уведомлений нет</div>
                )}
              </div>
            )}
          </div>
          <div className="current-user-badge">
            <strong>👤 {activeUser.name}</strong>
            <span>{activeUser.role === 'Руководитель' ? 'режим контроля руководителя' : roleWorkspaceHint(activeUser.role)}</span>
          </div>
          {isAdminPreviewAvailable && (
            <>
              <label style={{ display: 'grid', gap: '4px', fontSize: '12px', color: '#475569' }}>
                <span>Дивитись як:</span>
                <select
                  className="role-select"
                  value={currentRole}
                  onChange={(event) => {
                    const nextRole = event.target.value as Role;
                    setAdminPreviewRole(nextRole);
                    setAdminPreviewUserId('');
                    if (nextRole === 'Менеджер') setPage('dashboard');
                  }}
                  aria-label="Режим перегляду CRM за роллю"
                >
                  {adminPreviewRoles.map((role) => <option key={role} value={role}>{adminPreviewRoleLabel(role)}</option>)}
                </select>
              </label>
              <label style={{ display: 'grid', gap: '4px', fontSize: '12px', color: '#475569' }}>
                <span>Користувач:</span>
                <select
                  className="role-select"
                  value={adminPreviewUserId}
                  onChange={(event) => {
                    setAdminPreviewUserId(event.target.value);
                  }}
                  aria-label="Вибір співробітника для перегляду"
                >
                  <option value="">{currentRole === activeUser.role ? activeUser.name : 'Оберіть співробітника'}</option>
                  {previewUsers.map((user) => <option key={user.id} value={user.id}>{user.name}</option>)}
                </select>
              </label>
            </>
          )}
          {isOwnerControlView && (
            <button
              className="return-owner-button"
              type="button"
              onClick={() => {
                setActiveUserId(sessionUser.id);
                prependActionLog({ id: uid('LOG'), date: today, user: sessionUser.name, role: sessionUser.role, action: 'Повернення до керівника', entity: 'Робочий стіл', comment: 'Керівник повернувся зі службового перегляду ролі.' });
                stayOnCurrentPage();
              }}
            >
              До керівника
            </button>
          )}
          {canSwitchUsers && (
            <select
              className="role-select"
              value={activeUserId}
              onChange={(event) => {
                const nextUser = users.find((user) => user.id === event.target.value) ?? sessionUser;
                setActiveUserId(nextUser.id);
                prependActionLog({ id: uid('LOG'), date: today, user: sessionUser.name, role: sessionUser.role, action: 'Контроль ролі', entity: 'Робочий стіл', comment: `Керівник відкрив робочий стіл ролі ${nextUser.role}: ${nextUser.name}.` });
              }}
              aria-label="Перемикання ролі для контролю керівником"
            >
              {users.map((user) => <option key={user.id} value={user.id}>{user.name} · {user.role}</option>)}
            </select>
          )}
          <button className="icon-button" style={{ marginLeft: 'auto' }} type="button" onClick={logoutEmployee} aria-label="Вийти" title="Вийти">
            <LogOut size={18} />
          </button>
        </header>

        <nav className={`workspace-nav${showMenu ? ' workspace-nav-open' : ''}`} aria-label="Основна навігація">
          {visibleNavItems.map((item) => (
            <button
              type="button"
              key={item.id}
              className={page === item.id ? 'nav-item nav-item-active' : 'nav-item'}
              onClick={() => {
                setPage(item.id);
                setShowMenu(false);
              }}
            >
              {item.icon}
              {navLabelForRole(item.id, viewUser.role)}
            </button>
          ))}
        </nav>

        <ToastStack toasts={toasts} onDismiss={(id) => setToasts((current) => current.filter((toast) => toast.id !== id))} />

        {!canViewPage(page) && <AccessDenied activeUser={viewUser} page={page} />}
        {canViewPage(page) && page === 'dashboard' && viewUser.role === 'Менеджер' && (
          <OrdersPage
            orders={repairOrders}
            allOrders={orders}
            selectedOrder={selectedRepairOrder}
            products={products}
            selectedOrderId={selectedOrderId}
            setSelectedOrderId={setSelectedOrderId}
            selectedProductId={selectedProductId}
            setSelectedProductId={setSelectedProductId}
            qty={qty}
            setQty={setQty}
            quickPhone={quickPhone}
            quickClientDebtWarning={quickClientDebtWarning}
            quickClientName={quickClientName}
            quickDevice={quickDevice}
            quickSerial={quickSerial}
            quickProblem={quickProblem}
            quickAppearance={quickAppearance}
            quickEstimatedAmount={quickEstimatedAmount}
            quickEngineerId={quickEngineerId}
            quickContractId={quickContractId}
            quickLocationCode={quickLocationCode}
            quickComment={quickComment}
            users={users}
            contracts={contracts}
            contractActs={contractActs}
            bankImportItems={bankImportItems}
            setQuickPhone={handleQuickPhoneChange}
            setQuickClientName={setQuickClientName}
            setQuickDevice={setQuickDevice}
            setQuickSerial={setQuickSerial}
            setQuickProblem={setQuickProblem}
            setQuickAppearance={setQuickAppearance}
            setQuickEstimatedAmount={setQuickEstimatedAmount}
            setQuickEngineerId={setQuickEngineerId}
            setQuickContractId={setQuickContractId}
            setQuickLocationCode={setQuickLocationCode}
            setQuickComment={setQuickComment}
            createQuickOrder={createQuickOrder}
            patchSimpleManagerOrder={patchSimpleManagerOrder}
            updateSimpleManagerOrderStatus={updateSimpleManagerOrderStatus}
            editSimpleManagerOrder={editSimpleManagerOrder}
            addSimpleManagerOrderPart={addSimpleManagerOrderPart}
            removeSimpleManagerOrderPart={removeSimpleManagerOrderPart}
            acceptSimpleManagerPayment={acceptSimpleManagerPayment}
            confirmSimpleManagerPayment={confirmBankPayment}
            accountSimpleManagerOrderToContract={accountSimpleManagerOrderToContract}
            customerList={customerList}
            addPartToRepair={addPartToRepair}
            orderPart={orderPart}
            reserveArrived={reserveArrived}
            issueToEngineer={issueToEngineer}
            markInstalled={markInstalled}
            returnServicePart={returnServicePart}
            addOrderPayment={addOrderPayment}
            changeOrderStatus={changeOrderStatus}
            acceptOrderWork={acceptOrderWork}
            returnOrderToCellReady={returnOrderToCellReady}
            reassignEngineer={reassignEngineer}
            closeOrder={closeOrder}
            issueReadyOrder={issueReadyOrder}
            oneClickManagerIssue={oneClickManagerIssue}
            transferOrderToBas={transferOrderToBas}
            printDocument={printDocument}
            createServiceOrderDocument={createServiceOrderDocument}
            printServiceOrderDocument={printServiceOrderDocument}
            documents={documents}
            taxInvoices={taxInvoices}
            notifications={clientNotifications.filter((item) => item.orderId === selectedRepairOrder.id)}
            allNotifications={clientNotifications}
            orderUnits={orderUnits}
            warehouseLocations={warehouseLocations}
            orderMovementLogs={orderMovementLogs}
            moveOrder={moveOrder}
            sendClientNotification={enqueueClientNotification}
            createNotificationDraft={createNotificationDraft}
            approval={repairApprovals.find((item) => item.orderId === selectedRepairOrder.id)}
            sendRepairApproval={sendRepairApproval}
            recordApprovalResponse={recordApprovalResponse}
            markApprovalNoAnswer={markApprovalNoAnswer}
            logRiskConfirmation={logRiskConfirmation}
            ensureOrderDocumentRecord={ensureOrderDocumentRecord}
            logOrderDocumentPrint={logOrderDocumentPrint}
            signOrderAct={signOrderAct}
            createTaxInvoiceForOrder={createTaxInvoiceForOrder}
            registerTaxInvoice={registerTaxInvoice}
            markEngineerWorkCompleted={markEngineerWorkCompleted}
            cancelOrder={cancelOrder}
            refundOrder={refundOrder}
            cancelOrderAct={cancelOrderAct}
            reopenOrder={reopenOrder}
            openClientRecord={(phone, search) => {
              setGlobalFocusedClientPhone(phone);
              setGlobalClientSearch(search ?? '');
              stayOnCurrentPage();
            }}
            orderVersions={orderVersions}
            suggestedDocumentAction={suggestedDocumentAction}
            clearSuggestedDocumentAction={() => setSuggestedDocumentAction(null)}
            dashboardFocus={dashboardFocus}
            clearDashboardFocus={() => setDashboardFocus(null)}
            notifyUser={setNotice}
            canDo={canDo}
            showCost={canDo('cost.view')}
            activeUser={viewUser}
          />
        )}
        {canViewPage(page) && page === 'dashboard' && viewUser.role !== 'Менеджер' && (
          <Dashboard
            analytics={analytics}
            orders={orders}
            sales={sales}
            simplePayments={simplePayments}
            products={products}
            documents={documents}
            actionLogs={actionLogs}
            notifications={clientNotifications}
            repairApprovals={repairApprovals}
            requirements={requirements}
            purchases={purchases}
            taxInvoices={taxInvoices}
            setPage={setPage}
            activeUser={viewUser}
            users={users}
            quickPhone={quickPhone}
            quickClientDebtWarning={quickClientDebtWarning}
            quickClientName={quickClientName}
            quickDevice={quickDevice}
            quickSerial={quickSerial}
            quickProblem={quickProblem}
            quickAppearance={quickAppearance}
            quickEstimatedAmount={quickEstimatedAmount}
            quickEngineerId={quickEngineerId}
            quickContractId={quickContractId}
            quickLocationCode={quickLocationCode}
            quickComment={quickComment}
            customerList={customerList}
            setQuickPhone={handleQuickPhoneChange}
            setQuickClientName={setQuickClientName}
            setQuickDevice={setQuickDevice}
            setQuickSerial={setQuickSerial}
            setQuickProblem={setQuickProblem}
            setQuickAppearance={setQuickAppearance}
            setQuickEstimatedAmount={setQuickEstimatedAmount}
            setQuickEngineerId={setQuickEngineerId}
            setQuickContractId={setQuickContractId}
            setQuickLocationCode={setQuickLocationCode}
            setQuickComment={setQuickComment}
            createQuickOrder={createQuickOrder}
            warehouseLocations={warehouseLocations}
            contracts={contracts}
            setSelectedOrderId={setSelectedOrderId}
            sendClientNotification={enqueueClientNotification}
            setDashboardFocus={setDashboardFocus}
            recordManagerCall={recordManagerCall}
          />
        )}
        {canViewPage(page) && page === 'inbox' && (
          <InboxPage
            activeUser={viewUser}
            orders={orders}
            products={products}
            requirements={requirements}
            purchases={purchases}
            bankImportItems={bankImportItems}
            setPage={setPage}
            setSelectedOrderId={setSelectedOrderId}
            setDashboardFocus={setDashboardFocus}
          />
        )}
        {canViewPage(page) && page === 'employee-control' && (
          <EmployeeControlPage
            orders={orders}
            users={users}
          />
        )}
        {canViewPage(page) && page === 'bank-import' && (
          <BankImportPage
            items={bankImportItems}
            draft={bankImportDraft}
            allCandidates={buildBankImportCandidates()}
            bankAccounts={companyBankAccounts}
            onImport={handleBankStatementImport}
            onConfirmReview={confirmBankImportReview}
            onFinalizeMapping={finalizeMappedBankImport}
          />
        )}
        {canViewPage(page) && page === 'contracts' && (
          <ContractsPage
            contracts={contracts}
            contractActs={contractActs}
            orders={orders}
            documents={documents}
            createContract={createContract}
            createContractAct={createContractAct}
            createInvoiceForContractAct={createInvoiceForContractAct}
            closeContract={closeContract}
            readOnly={viewUser.role === 'Руководитель'}
          />
        )}
        {canViewPage(page) && page === 'sales' && (
          <SalesPage
            sales={sales}
            selectedSale={selectedSale}
            products={products}
            selectedSaleId={selectedSaleId}
            setSelectedSaleId={setSelectedSaleId}
            selectedSaleProductId={selectedSaleProductId}
            setSelectedSaleProductId={setSelectedSaleProductId}
            saleQty={saleQty}
            setSaleQty={setSaleQty}
            addSaleItem={addSaleItem}
            reserveSale={reserveSale}
            acceptSalePayment={acceptSalePayment}
            issueSale={issueSale}
            returnSale={returnSale}
            cancelSale={cancelSale}
            printDocument={printDocument}
            canDo={canDo}
            showCost={canDo('cost.view')}
            showProfit={canDo('profit.view')}
          />
        )}
        {canViewPage(page) && page === 'orders' && (
          <OrdersPage
            orders={repairOrders}
            allOrders={orders}
            selectedOrder={selectedRepairOrder}
            products={products}
            selectedOrderId={selectedOrderId}
            setSelectedOrderId={setSelectedOrderId}
            selectedProductId={selectedProductId}
            setSelectedProductId={setSelectedProductId}
            qty={qty}
            setQty={setQty}
            quickPhone={quickPhone}
            quickClientDebtWarning={quickClientDebtWarning}
            quickClientName={quickClientName}
            quickDevice={quickDevice}
            quickSerial={quickSerial}
            quickProblem={quickProblem}
            quickAppearance={quickAppearance}
            quickEstimatedAmount={quickEstimatedAmount}
            quickEngineerId={quickEngineerId}
            quickContractId={quickContractId}
            quickLocationCode={quickLocationCode}
            quickComment={quickComment}
            users={users}
            contracts={contracts}
            contractActs={contractActs}
            bankImportItems={bankImportItems}
            setQuickPhone={handleQuickPhoneChange}
            setQuickClientName={setQuickClientName}
            setQuickDevice={setQuickDevice}
            setQuickSerial={setQuickSerial}
            setQuickProblem={setQuickProblem}
            setQuickAppearance={setQuickAppearance}
            setQuickEstimatedAmount={setQuickEstimatedAmount}
            setQuickEngineerId={setQuickEngineerId}
            setQuickContractId={setQuickContractId}
            setQuickLocationCode={setQuickLocationCode}
            setQuickComment={setQuickComment}
            createQuickOrder={createQuickOrder}
            patchSimpleManagerOrder={patchSimpleManagerOrder}
            updateSimpleManagerOrderStatus={updateSimpleManagerOrderStatus}
            editSimpleManagerOrder={editSimpleManagerOrder}
            addSimpleManagerOrderPart={addSimpleManagerOrderPart}
            removeSimpleManagerOrderPart={removeSimpleManagerOrderPart}
            acceptSimpleManagerPayment={acceptSimpleManagerPayment}
            confirmSimpleManagerPayment={confirmBankPayment}
            accountSimpleManagerOrderToContract={accountSimpleManagerOrderToContract}
            customerList={customerList}
            addPartToRepair={addPartToRepair}
            orderPart={orderPart}
            reserveArrived={reserveArrived}
            issueToEngineer={issueToEngineer}
            markInstalled={markInstalled}
            returnServicePart={returnServicePart}
            addOrderPayment={addOrderPayment}
            changeOrderStatus={changeOrderStatus}
            acceptOrderWork={acceptOrderWork}
            returnOrderToCellReady={returnOrderToCellReady}
            reassignEngineer={reassignEngineer}
            closeOrder={closeOrder}
            issueReadyOrder={issueReadyOrder}
            oneClickManagerIssue={oneClickManagerIssue}
            transferOrderToBas={transferOrderToBas}
            printDocument={printDocument}
            createServiceOrderDocument={createServiceOrderDocument}
            printServiceOrderDocument={printServiceOrderDocument}
            documents={documents}
            taxInvoices={taxInvoices}
            notifications={clientNotifications.filter((item) => item.orderId === selectedRepairOrder.id)}
            allNotifications={clientNotifications}
            orderUnits={orderUnits}
            warehouseLocations={warehouseLocations}
            orderMovementLogs={orderMovementLogs}
            moveOrder={moveOrder}
            sendClientNotification={enqueueClientNotification}
            createNotificationDraft={createNotificationDraft}
            approval={repairApprovals.find((item) => item.orderId === selectedRepairOrder.id)}
            sendRepairApproval={sendRepairApproval}
            recordApprovalResponse={recordApprovalResponse}
            markApprovalNoAnswer={markApprovalNoAnswer}
            logRiskConfirmation={logRiskConfirmation}
            ensureOrderDocumentRecord={ensureOrderDocumentRecord}
            logOrderDocumentPrint={logOrderDocumentPrint}
            signOrderAct={signOrderAct}
            createTaxInvoiceForOrder={createTaxInvoiceForOrder}
            registerTaxInvoice={registerTaxInvoice}
            markEngineerWorkCompleted={markEngineerWorkCompleted}
            cancelOrder={cancelOrder}
            refundOrder={refundOrder}
            cancelOrderAct={cancelOrderAct}
            reopenOrder={reopenOrder}
            openClientRecord={(phone, search) => {
              setGlobalFocusedClientPhone(phone);
              setGlobalClientSearch(search ?? '');
              stayOnCurrentPage();
            }}
            orderVersions={orderVersions}
            suggestedDocumentAction={suggestedDocumentAction}
            clearSuggestedDocumentAction={() => setSuggestedDocumentAction(null)}
            dashboardFocus={dashboardFocus}
            clearDashboardFocus={() => setDashboardFocus(null)}
            notifyUser={setNotice}
            canDo={canDo}
            showCost={canDo('cost.view')}
            activeUser={viewUser}
          />
        )}
        {canViewPage(page) && page === 'my-orders' && (
          <OrdersPage
            orders={myOrders}
            allOrders={orders}
            allRoleOrders={viewUser.role === 'Менеджер' ? managerOwnOrders : engineerOwnOrders}
            selectedOrder={selectedMyOrder}
            products={products}
            selectedOrderId={selectedMyOrder.id}
            setSelectedOrderId={setSelectedOrderId}
            selectedProductId={selectedProductId}
            setSelectedProductId={setSelectedProductId}
            qty={qty}
            setQty={setQty}
            quickPhone={quickPhone}
            quickClientDebtWarning={quickClientDebtWarning}
            quickClientName={quickClientName}
            quickDevice={quickDevice}
            quickSerial={quickSerial}
            quickProblem={quickProblem}
            quickAppearance={quickAppearance}
            quickEstimatedAmount={quickEstimatedAmount}
            quickEngineerId={quickEngineerId}
            quickContractId={quickContractId}
            quickLocationCode={quickLocationCode}
            quickComment={quickComment}
            users={users}
            contracts={contracts}
            contractActs={contractActs}
            bankImportItems={bankImportItems}
            setQuickPhone={handleQuickPhoneChange}
            setQuickClientName={setQuickClientName}
            setQuickDevice={setQuickDevice}
            setQuickSerial={setQuickSerial}
            setQuickProblem={setQuickProblem}
            setQuickAppearance={setQuickAppearance}
            setQuickEstimatedAmount={setQuickEstimatedAmount}
            setQuickEngineerId={setQuickEngineerId}
            setQuickContractId={setQuickContractId}
            setQuickLocationCode={setQuickLocationCode}
            setQuickComment={setQuickComment}
            createQuickOrder={createQuickOrder}
            patchSimpleManagerOrder={patchSimpleManagerOrder}
            updateSimpleManagerOrderStatus={updateSimpleManagerOrderStatus}
            editSimpleManagerOrder={editSimpleManagerOrder}
            addSimpleManagerOrderPart={addSimpleManagerOrderPart}
            removeSimpleManagerOrderPart={removeSimpleManagerOrderPart}
            acceptSimpleManagerPayment={acceptSimpleManagerPayment}
            confirmSimpleManagerPayment={confirmBankPayment}
            accountSimpleManagerOrderToContract={accountSimpleManagerOrderToContract}
            customerList={customerList}
            addPartToRepair={addPartToRepair}
            orderPart={orderPart}
            reserveArrived={reserveArrived}
            issueToEngineer={issueToEngineer}
            markInstalled={markInstalled}
            returnServicePart={returnServicePart}
            addOrderPayment={addOrderPayment}
            issueReadyOrder={issueReadyOrder}
            transferOrderToBas={transferOrderToBas}
            closeOrder={closeOrder}
            printDocument={printDocument}
            createServiceOrderDocument={createServiceOrderDocument}
            printServiceOrderDocument={printServiceOrderDocument}
            documents={documents}
            taxInvoices={taxInvoices}
            notifications={clientNotifications.filter((item) => item.orderId === selectedMyOrder.id)}
            allNotifications={clientNotifications}
            orderUnits={orderUnits}
            warehouseLocations={warehouseLocations}
            orderMovementLogs={orderMovementLogs}
            moveOrder={moveOrder}
            sendClientNotification={enqueueClientNotification}
            createNotificationDraft={createNotificationDraft}
            approval={repairApprovals.find((item) => item.orderId === selectedMyOrder.id)}
            sendRepairApproval={sendRepairApproval}
            recordApprovalResponse={recordApprovalResponse}
            markApprovalNoAnswer={markApprovalNoAnswer}
            logRiskConfirmation={logRiskConfirmation}
            ensureOrderDocumentRecord={ensureOrderDocumentRecord}
            logOrderDocumentPrint={logOrderDocumentPrint}
            signOrderAct={signOrderAct}
            createTaxInvoiceForOrder={createTaxInvoiceForOrder}
            registerTaxInvoice={registerTaxInvoice}
            markEngineerWorkCompleted={markEngineerWorkCompleted}
            cancelOrder={cancelOrder}
            refundOrder={refundOrder}
            cancelOrderAct={cancelOrderAct}
            reopenOrder={reopenOrder}
            openClientRecord={(phone, search) => {
              setGlobalFocusedClientPhone(phone);
              setGlobalClientSearch(search ?? '');
              stayOnCurrentPage();
            }}
            orderVersions={orderVersions}
            suggestedDocumentAction={suggestedDocumentAction}
            clearSuggestedDocumentAction={() => setSuggestedDocumentAction(null)}
            dashboardFocus={dashboardFocus}
            clearDashboardFocus={() => setDashboardFocus(null)}
            notifyUser={setNotice}
            canDo={canDo}
            showCost={canDo('cost.view')}
            activeUser={viewUser}
            acceptOrderWork={acceptOrderWork}
            returnOrderToCellReady={returnOrderToCellReady}
            changeOrderStatus={changeOrderStatus}
            reassignEngineer={reassignEngineer}
            oneClickManagerIssue={oneClickManagerIssue}
          />
        )}
        {canViewPage(page) && page === 'parts' && <PartsPage products={products} requirements={requirements} receipts={receipts} movements={movements} productName={productName} showCost={canDo('cost.view')} receiveStockIntake={receiveStockIntake} canDo={canDo} onImportProducts={handleProductsImport} />}
        {canViewPage(page) && page === 'purchases' && (
          <PurchasesPage
            purchases={purchases}
            requirements={requirements}
            products={products}
            productName={productName}
            receivePurchase={receivePurchase}
            createManualPurchaseRequest={createManualPurchaseRequest}
            createPurchaseFromRequirement={createPurchaseFromRequirement}
            updatePurchaseProcurementMeta={updatePurchaseProcurementMeta}
            printDocument={printDocument}
            canDo={canDo}
          />
        )}
        {canViewPage(page) && page === 'cash' && <CashPage orders={orders} sales={sales} cashShift={cashShift} activeUser={viewUser} confirmBankPayment={confirmBankPayment} openCashShift={openCashShift} closeCashShift={closeCashShift} cancelPayment={cancelPayment} refundPayment={refundPayment} />}
        {canViewPage(page) && page === 'returns' && <ReturnsPage sales={sales} orders={orders} products={products} productName={productName} />}
        {canViewPage(page) && page === 'documents' && <DocumentsPage documents={documents} templates={documentTemplates} />}
        {canViewPage(page) && page === 'tax-invoices' && <TaxInvoicesPage invoices={taxInvoices} orders={orders} createTaxInvoiceForOrder={createTaxInvoiceForOrder} registerTaxInvoice={registerTaxInvoice} />}
        {canViewPage(page) && page === 'bas-exchange' && <BasExchangePage items={basExchangeItems} mappings={basMappings} actionLogs={actionLogs} />}
        {canViewPage(page) && page === 'problem-clients' && <ProblemClientsPage orders={orders} documents={documents} setPage={setPage} setSelectedOrderId={setSelectedOrderId} />}
        {canViewPage(page) && page === 'order-units' && <OrderUnitsPage units={orderUnits} activeUser={viewUser} />}
        {canViewPage(page) && page === 'storage' && <StoragePage orders={orders} locations={warehouseLocations} movementLogs={orderMovementLogs} />}
        {canViewPage(page) && page === 'movements' && <MovementsPage movements={movements} productName={productName} />}
        {canViewPage(page) && page === 'clients' && (
          <ClientsPage
            clients={customerList}
            orders={orders}
            simplePayments={simplePayments}
            contracts={contracts}
            contractActs={contractActs}
            onImportClients={handleClientsImport}
            createContractFromOrders={createContractFromOrders}
            initialSearch={globalClientSearch}
            focusedClientPhone={globalFocusedClientPhone}
            onFocusedClientPhoneChange={setGlobalFocusedClientPhone}
          />
        )}
        {canViewPage(page) && page === 'finance' && (
          <FinancePage
            orders={orders}
            products={products}
            receipts={receipts}
            movements={movements}
            contracts={contracts}
            contractActs={contractActs}
            bankImportItems={bankImportItems}
            simplePayments={simplePayments}
            users={users}
            onOpenClients={() => stayOnCurrentPage()}
          />
        )}
        {canViewPage(page) && page === 'reports' && <ReportsPage orders={orders} sales={sales} purchases={purchases} receipts={receipts} movements={movements} products={products} actionLogs={actionLogs} cashShift={cashShift} canDo={canDo} exportAccounting={exportAccounting} />}
        {canViewPage(page) && page === 'payroll' && <PayrollPage orders={orders} activeUser={viewUser} />}
        {canViewPage(page) && page === 'acceptance' && <AcceptanceChecklistPage />}
        {canViewPage(page) && page === 'team' && <TeamPage users={users} setUsers={setUsers} activeUser={viewUser} />}
        {canViewPage(page) && page === 'settings' && (
          <SettingsPage
            actionLogs={actionLogs}
            activeUser={viewUser}
            backups={backups}
            createManualBackup={() => {
              const snapshot = createBackupSnapshot('manual', 'Ручна копія CRM', activeUser.name);
              logAction('Ручний backup', snapshot.id, 'Користувач створив резервну копію вручну.');
              setNotice('Резервну копію створено.');
            }}
            restoreBackup={restoreBackupSnapshot}
            exportBackup={exportBackupSnapshot}
            exportCurrentLiveData={exportCurrentLiveData}
            importBackupFile={importBackupFile}
          />
        )}
        <footer style={{ marginTop: '24px', padding: '16px 0 8px', color: '#64748b', fontSize: '14px' }}>
          Контакт: {companySettings.mainEmail}
        </footer>
      </main>
    </div>
  );
}

function PageTitle({ eyebrow, title, text }: { eyebrow: string; title: string; text: string }) {
  return (
    <section className="page-title">
      <span>{eyebrow}</span>
      <h1>{title}</h1>
      <p>{text}</p>
    </section>
  );
}

function ToastStack({ toasts, onDismiss }: { toasts: AppToast[]; onDismiss: (id: string) => void }) {
  if (toasts.length === 0) return null;
  return (
    <div className="toast-stack" aria-live="polite" aria-atomic="true">
      {toasts.map((toast) => (
        <div key={toast.id} className={`toast toast-${toast.tone}`}>
          <span>{toast.message}</span>
          <button type="button" onClick={() => onDismiss(toast.id)} aria-label="Закрити повідомлення">×</button>
        </div>
      ))}
    </div>
  );
}

function InboxPage({
  activeUser,
  orders,
  products,
  requirements,
  purchases,
  bankImportItems,
  setPage,
  setSelectedOrderId,
  setDashboardFocus,
}: {
  activeUser: User;
  orders: ServiceOrder[];
  products: Product[];
  requirements: PartRequirement[];
  purchases: PurchaseOrder[];
  bankImportItems: BankImportItem[];
  setPage: (page: Page) => void;
  setSelectedOrderId: (id: string) => void;
  setDashboardFocus: (focus: { orderId: string; target: DashboardFocusTarget } | null) => void;
}) {
  type InboxTone = 'urgent' | 'today' | 'later';
  type InboxTask = {
    id: string;
    tone: InboxTone;
    text: string;
    actionLabel: string;
    onAction: () => void;
  };

  const openOrder = (orderId: string, target?: DashboardFocusTarget) => {
    if (target) setDashboardFocus({ orderId, target });
    setSelectedOrderId(orderId);
    setPage('orders');
  };

  const tasks: InboxTask[] = [];

  if (activeUser.role === 'Менеджер') {
    orders
      .filter((order) => !order.engineer)
      .slice(0, 20)
      .forEach((order) => {
        tasks.push({
          id: `no-engineer-${order.id}`,
          tone: hoursSince(order.statusChangedAt) >= 24 ? 'urgent' : 'today',
          text: `${order.id} без інженера`,
          actionLabel: 'Призначити',
          onAction: () => openOrder(order.id, 'issue'),
        });
      });

    orders
      .filter((order) => !['Видано', 'Закрито', 'Скасовано'].includes(order.status) && hoursSince(order.statusChangedAt) >= 48)
      .sort((a, b) => hoursSince(b.statusChangedAt) - hoursSince(a.statusChangedAt))
      .slice(0, 20)
      .forEach((order) => {
        tasks.push({
          id: `stuck-${order.id}`,
          tone: 'urgent',
          text: `${order.id} без руху ${hoursSince(order.statusChangedAt)} год`,
          actionLabel: 'Відкрити',
          onAction: () => openOrder(order.id, 'issue'),
        });
      });

    orders
      .filter((order) => orderTotals(order).debt > 0)
      .sort((a, b) => daysSince(b.debtSince ?? b.intakeDate) - daysSince(a.debtSince ?? a.intakeDate))
      .slice(0, 20)
      .forEach((order) => {
        tasks.push({
          id: `debt-${order.id}`,
          tone: simpleRepairStatus(order.status) === 'Готово' || order.status === 'Видано' || daysSince(order.debtSince ?? order.intakeDate) >= 2 ? 'urgent' : 'today',
          text: `${order.id} є борг ${money(orderTotals(order).debt)}`,
          actionLabel: 'Прийняти оплату',
          onAction: () => openOrder(order.id, 'payment'),
        });
      });

    orders
      .filter((order) => simpleRepairStatus(order.status) === 'Готово' && orderTotals(order).debt <= 0)
      .slice(0, 20)
      .forEach((order) => {
        tasks.push({
          id: `issue-${order.id}`,
          tone: 'today',
          text: `${order.id} готово до видачі`,
          actionLabel: 'Видати',
          onAction: () => openOrder(order.id, 'issue'),
        });
      });

    if (bankImportItems.some((item) => item.status === 'review')) {
      tasks.push({
        id: 'payments-review',
        tone: 'urgent',
        text: `Платежі на перевірці: ${bankImportItems.filter((item) => item.status === 'review').length}`,
        actionLabel: 'Перевірити',
        onAction: () => setPage('bank-import'),
      });
    }
  }

  if (activeUser.role === 'Інженер') {
    orders
      .filter((order) => !order.engineerAcceptedAt)
      .slice(0, 20)
      .forEach((order) => {
        tasks.push({
          id: `start-${order.id}`,
          tone: 'today',
          text: `${order.id} взяти в роботу`,
          actionLabel: 'Почати',
          onAction: () => setPage('my-orders'),
        });
      });

    orders
      .filter((order) => ['На діагностиці', 'В ремонті', 'На тестуванні'].includes(order.status))
      .sort((a, b) => daysSince(b.statusChangedAt) - daysSince(a.statusChangedAt))
      .slice(0, 20)
      .forEach((order) => {
        tasks.push({
          id: `finish-${order.id}`,
          tone: hoursSince(order.statusChangedAt) >= ENGINEER_OVERDUE_HOURS ? 'urgent' : 'later',
          text: `${order.id} завершити роботу`,
          actionLabel: 'Готово',
          onAction: () => setPage('my-orders'),
        });
      });
  }

  if (activeUser.role === 'Закупник' || activeUser.role === 'Комірник') {
    products
      .filter((product) => available(product) < Math.max(product.min, 1))
      .sort((a, b) => (available(a) - a.min) - (available(b) - b.min))
      .slice(0, 20)
      .forEach((product) => {
        const hasOpenPurchase = purchases.some((purchase) =>
          ['Нова', 'В роботі', 'Замовлено', 'В дорозі'].includes(purchase.status)
          && purchase.items.some((item) => item.productId === product.id),
        );
        tasks.push({
          id: `stock-${product.id}`,
          tone: available(product) < Math.max(product.min, 1) ? 'urgent' : 'today',
          text: `${product.sku} нижче мінімуму`,
          actionLabel: hasOpenPurchase ? 'Відкрити закупку' : 'Створити закупку',
          onAction: () => setPage('purchases'),
        });
      });

    requirements
      .filter((item) => ['Замовлено', 'В дорозі'].includes(item.status))
      .slice(0, 20)
      .forEach((item) => {
        const product = products.find((entry) => entry.id === item.productId);
        tasks.push({
          id: `purchase-${item.id}`,
          tone: item.status === 'В дорозі' ? 'today' : 'later',
          text: `${product?.name ?? item.productId} ${item.status.toLowerCase()}`,
          actionLabel: 'Відкрити',
          onAction: () => setPage('purchases'),
        });
      });
  }

  const grouped = {
    urgent: tasks.filter((task) => task.tone === 'urgent'),
    today: tasks.filter((task) => task.tone === 'today'),
    later: tasks.filter((task) => task.tone === 'later'),
  };

  return (
    <div className="page-grid">
      <PageTitle eyebrow="Inbox" title="Центр задач" text="Одна задача, одна дія, одна кнопка." />
      <section className="panel inbox-panel">
        {([
          { key: 'urgent', title: 'Терміново', tone: 'urgent' },
          { key: 'today', title: 'Сьогодні', tone: 'today' },
          { key: 'later', title: 'Інше', tone: 'later' },
        ] as const).map((group) => (
          <div key={group.key} className="inbox-group">
            <div className={`inbox-group-head inbox-group-head-${group.tone}`}>
              <strong>{group.title}</strong>
              <span>{grouped[group.key].length}</span>
            </div>
            <div className="inbox-list">
              {grouped[group.key].map((task) => (
                <div key={task.id} className="inbox-row">
                  <span>{task.text}</span>
                  <button type="button" onClick={task.onAction}>{task.actionLabel}</button>
                </div>
              ))}
              {grouped[group.key].length === 0 && <div className="empty-state">Немає задач.</div>}
            </div>
          </div>
        ))}
      </section>
    </div>
  );
}

function AccessDenied({ activeUser, page }: { activeUser: User; page: Page }) {
  const pageName = navLabelForRole(page, activeUser.role);
  return (
    <div className="page-grid">
      <PageTitle eyebrow="Доступ обмежено" title="Этот раздел скрыт для роли" text={`Користувач ${activeUser.name} має роль ${activeUser.role}. Розділ "${pageName}" не входить у мінімально необхідний доступ для цієї ролі.`} />
      <section className="panel">
        <div className="task-list">
          <Task icon={<ShieldCheck />} title="Принцип безпеки" text="Співробітник бачить тільки ті модулі й дії, які потрібні для його роботи." />
          <Task icon={<History />} title="Журнал" text="Спроби зміни ролі та критичні дії фіксуються в журналі." />
        </div>
      </section>
    </div>
  );
}

function MessageFeed({
  messages,
  activeUser,
  users,
  orders,
  onSend,
  onStatus,
}: {
  messages: InternalMessage[];
  activeUser: User;
  users: User[];
  orders: ServiceOrder[];
  onSend: (toUserId: string, orderId: string, text: string, dueAt: string, importance: InternalMessageImportance) => void;
  onStatus: (messageId: string, status: InternalMessageStatus) => void;
}) {
  const [toUserId, setToUserId] = useState(users.find((user) => user.role === 'Менеджер')?.id ?? users[0].id);
  const [orderId, setOrderId] = useState(orders[0]?.id ?? '');
  const [text, setText] = useState('');
  const [dueAt, setDueAt] = useState('20.04.2026 18:00');
  const [importance, setImportance] = useState<InternalMessageImportance>('Важлива');

  if (messages.length === 0) return null;

  return (
    <section className="message-feed" aria-label="Внутрішні повідомлення">
      <div className="message-feed-head">
        <div>
          <span>{activeUser.role === 'Руководитель' ? 'Центр повідомлень' : 'Мої задачі'}</span>
          <h2>{activeUser.role === 'Руководитель' ? 'Адресні повідомлення команди' : 'Адресні повідомлення для мене'}</h2>
        </div>
        <strong>{messages.filter((message) => message.status === 'Нове').length} нових</strong>
      </div>

      {activeUser.role === 'Руководитель' && (
        <form
          className="message-feed-form"
          onSubmit={(event) => {
            event.preventDefault();
            onSend(toUserId, orderId, text, dueAt, importance);
            setText('');
          }}
        >
          <select value={toUserId} onChange={(event) => setToUserId(event.target.value)} aria-label="Кому">
            {users.filter((user) => user.role !== 'Руководитель').map((user) => <option key={user.id} value={user.id}>{user.role} — {user.name}</option>)}
          </select>
          <select value={orderId} onChange={(event) => setOrderId(event.target.value)} aria-label="Замовлення">
            <option value="">Без замовлення</option>
            {orders.map((order) => <option key={order.id} value={order.id}>{order.id} — {order.client}</option>)}
          </select>
          <select value={importance} onChange={(event) => setImportance(event.target.value as InternalMessageImportance)} aria-label="Важливість">
            <option value="Звичайна">Звичайна</option>
            <option value="Важлива">Важлива</option>
            <option value="Критична">Критична</option>
          </select>
          <input value={dueAt} onChange={(event) => setDueAt(event.target.value)} aria-label="Строк" placeholder="Строк" />
          <textarea value={text} onChange={(event) => setText(event.target.value)} placeholder="Текст задачі або повідомлення" />
          <button type="submit" className="submit-button">Надіслати</button>
        </form>
      )}

      <div className="message-feed-list">
        {messages.slice(0, 5).map((message) => (
          <article className={`message-item importance-${message.importance.toLowerCase()}`} key={message.id}>
            <div>
              <strong>{message.toRole} — {message.toUser}</strong>
              <p>{message.text}</p>
              <span>{message.orderId ? `Замовлення: ${message.orderId}. ` : ''}Строк: {message.dueAt}. Від: {message.createdBy}</span>
            </div>
            <div className="message-item-actions">
              <span className={`status-pill ${message.status === 'Нове' ? 'danger' : message.status === 'Виконано' ? 'good' : ''}`}>{message.status}</span>
              <span>{message.importance}</span>
              {message.status === 'Нове' && <button type="button" onClick={() => onStatus(message.id, 'Прочитано')}>Прочитано</button>}
              {message.status !== 'Виконано' && <button type="button" onClick={() => onStatus(message.id, 'Виконано')}>Виконано</button>}
            </div>
          </article>
        ))}
      </div>
    </section>
  );
}

function QuickOrderIntake({
  phone,
  clientName,
  device,
  serial,
  problem,
  appearance,
  estimatedAmount,
  engineerId,
  contractId,
  locationCode,
  comment,
  managerName,
  users,
  contracts,
  customerList,
  warehouseLocations,
  onPhoneChange,
  onClientNameChange,
  onDeviceChange,
  onSerialChange,
  onProblemChange,
  onAppearanceChange,
  onEstimatedAmountChange,
  onEngineerChange,
  onContractChange,
  onLocationChange,
  onCommentChange,
  onSubmit,
}: {
  phone: string;
  clientName: string;
  device: string;
  serial: string;
  problem: string;
  appearance: string;
  estimatedAmount: string;
  engineerId: string;
  contractId: string;
  locationCode: string;
  comment: string;
  managerName: string;
  users: User[];
  contracts: ContractRecord[];
  customerList: ClientRecord[];
  warehouseLocations: WarehouseLocation[];
  onPhoneChange: (phone: string) => void;
  onClientNameChange: (name: string) => void;
  onDeviceChange: (device: string) => void;
  onSerialChange: (serial: string) => void;
  onProblemChange: (problem: string) => void;
  onAppearanceChange: (appearance: string) => void;
  onEstimatedAmountChange: (amount: string) => void;
  onEngineerChange: (id: string) => void;
  onContractChange: (id: string) => void;
  onLocationChange: (code: string) => void;
  onCommentChange: (comment: string) => void;
  onSubmit: () => void;
}) {
  const normalized = phone.replace(/\D/g, '');
  const foundClient = normalized.length >= 7
    ? customerList.find((client) => client.phone.replace(/\D/g, '').includes(normalized) || normalized.includes(client.phone.replace(/\D/g, '')))
    : undefined;
  const engineers = users.filter((user) => user.role === 'Інженер');
  const activeContracts = contracts.filter((contract) => contract.status === 'Активний');
  const repairLocations = warehouseLocations.filter((location) => location.zone === 'REPAIR');

  return (
    <section className="quick-intake-panel">
      <div className="panel-heading">
        <div>
          <h2>Створення замовлення ремонту</h2>
          <span>Основний робочий екран менеджера</span>
        </div>
        <strong>{foundClient ? `Клієнта знайдено: ${foundClient.orders} звернень` : 'Новий клієнт створиться автоматично'}</strong>
      </div>
      <form
        className="quick-intake-form"
        onSubmit={(event) => {
          event.preventDefault();
          onSubmit();
        }}
      >
        <label>
          Телефон клієнта
          <input value={phone} onChange={(event) => onPhoneChange(event.target.value)} placeholder="+380..." />
        </label>
        <label>
          Клієнт
          <input value={clientName} onChange={(event) => onClientNameChange(event.target.value)} placeholder="ПІБ або назва компанії" />
        </label>
        <label>
          Пристрій
          <input value={device} onChange={(event) => onDeviceChange(event.target.value)} placeholder="Ноутбук Lenovo, МФУ Canon..." />
        </label>
        <label>
          Серійний номер
          <input value={serial} onChange={(event) => onSerialChange(event.target.value)} placeholder="SN / IMEI / сервісний номер" />
        </label>
        <label>
          Проблема клієнта
          <input value={problem} onChange={(event) => onProblemChange(event.target.value)} placeholder="Не вмикається, не друкує..." />
        </label>
        <label>
          Зовнішній вигляд / комплектація
          <input value={appearance} onChange={(event) => onAppearanceChange(event.target.value)} placeholder="Зарядка, сумка, подряпини, без АКБ..." />
        </label>
        <label>
          Орієнтовна сума
          <input value={estimatedAmount} onChange={(event) => onEstimatedAmountChange(event.target.value)} placeholder="0" type="number" min={0} />
        </label>
        <label>
          Менеджер
          <input value={managerName} readOnly />
        </label>
        <label>
          Призначений інженер
          <select value={engineerId} onChange={(event) => onEngineerChange(event.target.value)}>
            {engineers.map((engineer) => <option key={engineer.id} value={engineer.id}>{engineer.name}</option>)}
          </select>
        </label>
        <label>
          Договір
          <select value={contractId} onChange={(event) => onContractChange(event.target.value)}>
            <option value="">Без договору</option>
            {activeContracts.map((contract) => <option key={contract.id} value={contract.id}>{contract.id} · {contract.client}</option>)}
          </select>
        </label>
        <label>
          Закріплена комірка
          <select value={locationCode} onChange={(event) => onLocationChange(event.target.value)}>
            <option value="">Автопідбір</option>
            {repairLocations.map((location) => <option key={location.id} value={location.code}>{location.code}</option>)}
          </select>
        </label>
        <label style={{ gridColumn: '1 / -1' }}>
          Коментар
          <input value={comment} onChange={(event) => onCommentChange(event.target.value)} placeholder="Коментар менеджера при прийомі" />
        </label>
        <button className="submit-button" type="submit">Прийняти замовлення</button>
      </form>
    </section>
  );
}

function Dashboard({
  analytics,
  orders,
  sales,
  simplePayments,
  products,
  documents,
  actionLogs,
  notifications,
  repairApprovals,
  requirements,
  purchases,
  taxInvoices,
  setPage,
  activeUser,
  users,
  quickPhone,
  quickClientName,
  quickDevice,
  quickSerial,
  quickProblem,
  quickAppearance,
  quickEstimatedAmount,
  quickEngineerId,
  quickContractId,
  quickLocationCode,
  quickComment,
  customerList,
  setQuickPhone,
  setQuickClientName,
  setQuickDevice,
  setQuickSerial,
  setQuickProblem,
  setQuickAppearance,
  setQuickEstimatedAmount,
  setQuickEngineerId,
  setQuickContractId,
  setQuickLocationCode,
  setQuickComment,
  createQuickOrder,
  warehouseLocations,
  contracts,
  setSelectedOrderId,
  sendClientNotification,
  setDashboardFocus,
  recordManagerCall,
}: {
  analytics: { stockValue: number; partsInRepairs: number; ordered: number; required: number; salesRevenue?: number; salesDebt?: number; serviceDebt?: number; salesProfit?: number; serviceProfit?: number };
  orders: ServiceOrder[];
  sales: Sale[];
  simplePayments: SimpleOrderPaymentRecord[];
  products: Product[];
  documents: PrintDocument[];
  actionLogs: ActionLog[];
  notifications: ClientNotification[];
  repairApprovals: RepairApproval[];
  requirements: PartRequirement[];
  purchases: PurchaseOrder[];
  taxInvoices: TaxInvoice[];
  setPage: (page: Page) => void;
  activeUser: User;
  users: User[];
  quickPhone: string;
  quickClientDebtWarning: string;
  quickClientName: string;
  quickDevice: string;
  quickSerial: string;
  quickProblem: string;
  quickAppearance: string;
  quickEstimatedAmount: string;
  quickEngineerId: string;
  quickContractId: string;
  quickLocationCode: string;
  quickComment: string;
  customerList: ClientRecord[];
  setQuickPhone: (phone: string) => void;
  setQuickClientName: (name: string) => void;
  setQuickDevice: (device: string) => void;
  setQuickSerial: (serial: string) => void;
  setQuickProblem: (problem: string) => void;
  setQuickAppearance: (appearance: string) => void;
  setQuickEstimatedAmount: (amount: string) => void;
  setQuickEngineerId: (id: string) => void;
  setQuickContractId: (id: string) => void;
  setQuickLocationCode: (code: string) => void;
  setQuickComment: (comment: string) => void;
  createQuickOrder: () => void;
  warehouseLocations: WarehouseLocation[];
  contracts: ContractRecord[];
  setSelectedOrderId: (id: string) => void;
  sendClientNotification: (order: ServiceOrder, event: NotificationEvent, preferredChannel?: NotificationChannel, manual?: boolean, messageOverride?: string) => void;
  setDashboardFocus: (focus: { orderId: string; target: DashboardFocusTarget } | null) => void;
  recordManagerCall: (order: ServiceOrder) => void;
}) {
  const [showOwnerLeakQueue, setShowOwnerLeakQueue] = useState(false);
  const newOrders = orders.filter((order) => order.status === 'Прийнято').length;
  const diagnostics = orders.filter((order) => order.status === 'На діагностиці').length;
  const waitingParts = orders.filter((order) => order.status === 'Очікує запчастину').length;
  const inRepair = orders.filter((order) => order.status === 'В ремонті').length;
  const ready = orders.filter((order) => ['Готовий до видачі', 'Не підлягає ремонту'].includes(order.status)).length;
  const issuedToday = orders.filter((order) => order.status === 'Видано').length;
  const stuckOrders = orders.filter((order) => daysSince(order.statusChangedAt) > 2);
  const minStock = products.filter((product) => available(product) <= product.min);
  const salesToday = sales.reduce((sum, sale) => sum + saleTotals(sale).paid, 0);
  const serviceRevenue = orders.reduce((sum, order) => sum + orderTotals(order).paid, 0);
  const monthRevenue = salesToday + serviceRevenue;
  const monthProfit = (analytics.salesProfit ?? 0) + (analytics.serviceProfit ?? 0);
  const debt = (analytics.salesDebt ?? 0) + (analytics.serviceDebt ?? 0);
  const activeRepairs = orders.filter((order) => !['Закрито', 'Скасовано', 'Видано', 'Не підлягає ремонту'].includes(order.status)).length;
  const averageCheck = sales.length + orders.length > 0 ? Math.round(monthRevenue / (sales.length + orders.length)) : 0;
  const returnsAmount = sales.reduce((sum, sale) => sum + sale.returns.reduce((inner, item) => inner + item.refund, 0), 0);
  const writeOffs = products.reduce((sum, product) => sum + product.installed * product.cost, 0);
  const overdueOrders = orders.filter((order) => daysSince(order.promisedDate) > 0 && !['Закрито', 'Видано', 'Скасовано'].includes(order.status));
  const problemOrders = [...stuckOrders, ...overdueOrders, ...orders.filter((order) => order.status === 'Очікує запчастину' && daysSince(order.statusChangedAt) > 2)].filter((order, index, list) => list.findIndex((item) => item.id === order.id) === index);
  const statusBars = [
    { label: 'Діагностика', value: diagnostics },
    { label: 'Ремонт', value: inRepair },
    { label: 'Запчастина', value: waitingParts },
    { label: 'Готово', value: ready },
    { label: 'Просрочені', value: overdueOrders.length },
  ];
  const revenueBars = [
    { label: 'Ремонт', value: serviceRevenue },
    { label: 'Продажі', value: analytics.salesRevenue ?? 0 },
    { label: 'Борги', value: debt },
    { label: 'Повернення', value: returnsAmount },
  ];
  const managerRows = users.filter((user) => ['Менеджер', 'Адміністратор'].includes(user.role)).map((user) => {
    const managerOrders = orders.filter((order) => order.manager === user.name);
    const managerSales = sales.filter((sale) => sale.manager === user.name);
    const revenue = managerSales.reduce((sum, sale) => sum + saleTotals(sale).paid, 0) + managerOrders.reduce((sum, order) => sum + orderTotals(order).paid, 0);
    const closed = managerOrders.filter((order) => ['Закрито', 'Видано'].includes(order.status)).length;
    const paymentsCount = managerOrders.reduce((sum, order) => sum + order.payments.length, 0) + managerSales.reduce((sum, sale) => sum + sale.payments.length, 0);
    const documentsCount = managerOrders.filter((order) => order.legalEntity && order.actIssuedAt && !order.actReturnedAt).length;
    const managerDebt = managerOrders.reduce((sum, order) => sum + orderTotals(order).debt, 0) + managerSales.reduce((sum, sale) => sum + saleTotals(sale).debt, 0);
    const avg = managerOrders.length + managerSales.length ? Math.round(revenue / (managerOrders.length + managerSales.length)) : 0;
    return { name: user.name, orders: managerOrders.length, closed, sales: managerSales.length, revenue, paymentsCount, avg, debt: managerDebt, documentsCount };
  }).filter((row) => row.orders || row.sales);
  const engineerRows = users.filter((user) => user.role === 'Інженер').map((user) => {
    const engineerOrders = orders.filter((order) => order.engineer === user.name);
    const finished = engineerOrders.filter((order) => ['Готовий до видачі', 'Не підлягає ремонту', 'Видано', 'Закрито'].includes(order.status)).length;
    const avgDays = engineerOrders.length ? Math.round(engineerOrders.reduce((sum, order) => sum + daysSince(order.intakeDate), 0) / engineerOrders.length) : 0;
    const stuck = engineerOrders.filter((order) => daysSince(order.statusChangedAt) > 2 && !['Закрито', 'Видано', 'Скасовано'].includes(order.status)).length;
    const repeated = engineerOrders.filter((order) => order.clientComment?.toLowerCase().includes('повтор')).length;
    const earning = engineerOrders.reduce((sum, order) => sum + order.works.reduce((inner, work) => inner + payrollForWork(order, work, payrollRules.find((rule) => rule.employee === (work.engineer ?? order.engineer))), 0), 0);
    const companyProfit = engineerOrders.reduce((sum, order) => sum + orderTotals(order).finalProfit, 0);
    return { name: user.name, orders: engineerOrders.length, active: engineerOrders.filter((order) => !['Закрито', 'Видано', 'Скасовано', 'Не підлягає ремонту'].includes(order.status)).length, finished, avgDays, stuck, repeated, earning, companyProfit };
  }).filter((row) => row.active || row.finished);
  const topProducts = products.map((product) => {
    const sold = sales.reduce((sum, sale) => sum + sale.items.filter((item) => item.productId === product.id).reduce((inner, item) => inner + item.qty, 0), 0);
    return { product, sold, availableQty: available(product) };
  }).sort((a, b) => b.sold - a.sold);
  const isEngineer = activeUser.role === 'Інженер';
  const isManager = activeUser.role === 'Менеджер' || activeUser.role === 'Адміністратор';
  const isAccountant = activeUser.role === 'Бухгалтер';
  const isStock = activeUser.permissions.canAccessWarehouse;
  const isOwner = activeUser.role === 'Руководитель' || activeUser.role === 'Адміністратор';
  const roleOrders = isEngineer ? orders.filter((order) => order.engineer === activeUser.name) : orders;
  const roleActiveRepairs = roleOrders.filter((order) => !['Закрито', 'Скасовано', 'Видано', 'Не підлягає ремонту'].includes(order.status));
  const roleReadyOrders = roleOrders.filter((order) => ['Готовий до видачі', 'Не підлягає ремонту', 'Очікує клієнта', 'Очікує оплати'].includes(order.status));
  const noPaymentOrders = orders.filter((order) => orderTotals(order).debt > 0);
  const noActOrders = orders.filter((order) => order.legalEntity && order.actIssuedAt && !order.actReturnedAt);
  const serviceProblems = buildServiceProblems(orders, documents);
  const roleServiceProblems = isManager && !isOwner
    ? serviceProblems.filter((problem) => problem.manager === activeUser.name || problem.responsible === activeUser.name)
    : serviceProblems;
  const criticalProblems = roleServiceProblems.filter((problem) => problem.level === 'Критично');
  const overdueTaxInvoices = taxInvoices.filter((invoice) => invoice.status === 'Помилка');
  const missingParts = requirements.filter((item) => ['Потрібно', 'До закупівлі'].includes(item.status));
  const delayedPurchases = purchases.filter((purchase) => purchase.status !== 'Прибуло' && daysSince(purchase.expectedAt) > 0);
  const accountantKpi = {
    readyTax: taxInvoices.filter((invoice) => invoice.status === 'Створено').length,
    overdueTax: overdueTaxInvoices.length,
    acts: noActOrders.length,
    unpaid: noPaymentOrders.length,
  };
  const stockKpi = {
    receipts: purchases.filter((purchase) => purchase.status === 'Прибуло' || purchase.status === 'Частково прибуло').length,
    stock: products.reduce((sum, product) => sum + available(product), 0),
    critical: minStock.length,
    delays: delayedPurchases.length,
  };
  const payrollDashboardRows = buildPayrollEmployeeSummaries(orders);
  const managerOrders = orders.filter((order) => order.manager === activeUser.name);
  const managerWaitResponseOrders = managerOrders
    .map((order) => {
      const approval = repairApprovals.find((item) => item.orderId === order.id && item.status === 'Очікує відповідь');
      if (!approval) return null;
      return { order, approval, ageDays: daysSince(approval.sentAt), ageHours: hoursSince(approval.sentAt) };
    })
    .filter((item): item is { order: ServiceOrder; approval: RepairApproval; ageDays: number; ageHours: number } => Boolean(item))
    .sort((a, b) => b.ageHours - a.ageHours);
  const managerReadyNotIssuedOrders = managerOrders
    .filter((order) => ['Готовий до видачі', 'Очікує оплати', 'Очікує клієнта'].includes(order.status))
    .map((order) => ({ order, ageDays: daysSince(order.returnedToCellAt ?? order.statusChangedAt) }))
    .sort((a, b) => b.ageDays - a.ageDays);
  const managerUnpaidOrders = managerOrders
    .filter((order) => {
      const totals = orderTotals(order);
      return totals.total > 0 && totals.paid <= 0 && !['Видано', 'Закрито', 'Скасовано'].includes(order.status);
    })
    .map((order) => ({ order, ageDays: daysSince(order.statusChangedAt) }))
    .sort((a, b) => b.ageDays - a.ageDays);
  const MANAGER_OVERDUE_DAYS = 5;
  const managerOverdueOrders = managerOrders
    .filter((order) => !['Видано', 'Закрито', 'Скасовано'].includes(order.status) && daysSince(order.intakeDate) > MANAGER_OVERDUE_DAYS)
    .map((order) => ({ order, ageDays: daysSince(order.intakeDate) }))
    .sort((a, b) => b.ageDays - a.ageDays);
  const todayKey = extractDayKey(today);
  const managerMoneyWaitOrders = managerOrders
    .filter((order) => ['Готовий до видачі', 'Очікує оплати', 'Очікує клієнта'].includes(order.status) && orderTotals(order).debt > 0)
    .map((order) => ({ order, debt: orderTotals(order).debt, ageDays: daysSince(order.returnedToCellAt ?? order.statusChangedAt) }))
    .sort((a, b) => b.ageDays - a.ageDays || b.debt - a.debt);
  const managerPartialPaidOrders = managerOrders
    .filter((order) => {
      const totals = orderTotals(order);
      return totals.paid > 0 && totals.debt > 0 && !['Видано', 'Закрито', 'Скасовано'].includes(order.status);
    })
    .map((order) => {
      const totals = orderTotals(order);
      const lastPayment = [...order.payments].sort((a, b) => parseDateTime(b.date)?.getTime() ?? 0 - (parseDateTime(a.date)?.getTime() ?? 0))[0];
      return { order, paid: totals.paid, debt: totals.debt, ageDays: daysSince(lastPayment?.date ?? order.statusChangedAt), lastPayment };
    })
    .sort((a, b) => b.ageDays - a.ageDays || b.debt - a.debt);
  const managerPaidTodayRows = managerOrders
    .flatMap((order) => order.payments
      .filter((payment) => extractDayKey(payment.date) === todayKey && payment.amount > 0)
      .map((payment) => ({ order, payment })))
    .sort((a, b) => (parseDateTime(b.payment.date)?.getTime() ?? 0) - (parseDateTime(a.payment.date)?.getTime() ?? 0));
  const managerPaidTodayTotal = managerPaidTodayRows.reduce((sum, row) => sum + row.payment.amount, 0);
  const managerProblemMoneyOrders = managerOrders
    .filter((order) => {
      const totals = orderTotals(order);
      const unpaidOverdue = totals.debt > 0 && totals.paid <= 0 && daysSince(order.statusChangedAt) > 3;
      const partialOverdue = totals.paid > 0 && totals.debt > 0 && daysSince((order.payments[order.payments.length - 1]?.date ?? order.statusChangedAt)) > 2;
      const noAnswer = managerWaitResponseOrders.some((item) => item.order.id === order.id);
      return unpaidOverdue || partialOverdue || noAnswer;
    })
    .map((order) => ({ order, debt: orderTotals(order).debt, ageDays: daysSince(order.statusChangedAt) }))
    .sort((a, b) => b.ageDays - a.ageDays || b.debt - a.debt);
  const managerMoneyToday = managerPaidTodayTotal;
  const managerMoneyDebt = managerOrders.reduce((sum, order) => sum + orderTotals(order).debt, 0);
  const managerMoneyPartial = managerPartialPaidOrders.reduce((sum, item) => sum + item.debt, 0);
  const managerMoneyOverdue = managerProblemMoneyOrders.reduce((sum, item) => sum + item.debt, 0);
  const OWNER_OVERDUE_DAYS = 5;
  const ownerOverdueOrders = orders
    .filter((order) => !['Видано', 'Закрито', 'Скасовано'].includes(order.status) && daysSince(order.intakeDate) > OWNER_OVERDUE_DAYS)
    .map((order) => ({ order, ageDays: daysSince(order.intakeDate), amount: Math.max(orderTotals(order).debt, orderTotals(order).total) }))
    .sort((a, b) => b.ageDays - a.ageDays || b.amount - a.amount);
  const ownerEngineerActiveCount = orders.filter((order) => order.locationStatus === 'У інженера').length;
  const ownerEngineerStuckOrders = orders
    .filter((order) => order.locationStatus === 'У інженера' && order.takenInRepairAt && hoursSince(order.takenInRepairAt) > 24)
    .map((order) => ({ order, ageHours: hoursSince(order.takenInRepairAt as string) }))
    .sort((a, b) => b.ageHours - a.ageHours || Math.max(orderTotals(b.order).debt, orderTotals(b.order).total) - Math.max(orderTotals(a.order).debt, orderTotals(a.order).total));
  const ownerUnpaidReadyOrders = orders
    .filter((order) => ['Готовий до видачі', 'Очікує оплати', 'Очікує клієнта'].includes(order.status) && orderTotals(order).debt > 0)
    .map((order) => ({ order, debt: orderTotals(order).debt, ageDays: daysSince(order.returnedToCellAt ?? order.statusChangedAt) }))
    .sort((a, b) => b.ageDays - a.ageDays || b.debt - a.debt);
  const ownerReadyNotPickedOrders = orders
    .filter((order) => ['Готовий до видачі', 'Очікує клієнта'].includes(order.status) && orderTotals(order).debt <= 0)
    .map((order) => ({ order, ageDays: daysSince(order.returnedToCellAt ?? order.statusChangedAt), amount: orderTotals(order).total }))
    .sort((a, b) => b.ageDays - a.ageDays || b.amount - a.amount);
  const ownerLeakDebtTotal = orders.reduce((sum, order) => sum + orderTotals(order).debt, 0);
  const ownerLeakQueue = [
    ...ownerOverdueOrders.map((item) => ({ type: 'Прострочено', priority: 4, order: item.order, ageDays: item.ageDays, amount: item.amount, action: 'issue' as DashboardFocusTarget })),
    ...ownerEngineerStuckOrders.map((item) => ({ type: 'У інженера', priority: 3, order: item.order, ageDays: Math.floor(item.ageHours / 24), amount: Math.max(orderTotals(item.order).debt, orderTotals(item.order).total), action: 'issue' as DashboardFocusTarget })),
    ...ownerUnpaidReadyOrders.map((item) => ({ type: 'Не оплачено', priority: 2, order: item.order, ageDays: item.ageDays, amount: item.debt, action: 'payment' as DashboardFocusTarget })),
    ...ownerReadyNotPickedOrders.map((item) => ({ type: 'Не забрали', priority: 1, order: item.order, ageDays: item.ageDays, amount: item.amount, action: 'issue' as DashboardFocusTarget })),
  ].sort((a, b) => b.priority - a.priority || b.ageDays - a.ageDays || b.amount - a.amount);
  const managerCashierRows = Array.from(
    managerPaidTodayRows.reduce((map, row) => {
      const acceptedBy = row.payment.acceptedBy;
      const current = map.get(acceptedBy) ?? 0;
      map.set(acceptedBy, current + row.payment.amount);
      return map;
    }, new Map<string, number>()),
  ).sort((a, b) => b[1] - a[1]);
  const managerUrgencyStyle = (days: number) => (
    days > 3
      ? { borderColor: '#dc2626', boxShadow: '0 0 0 2px rgba(220, 38, 38, 0.12)' }
      : days > 1
        ? { borderColor: '#f59e0b', boxShadow: '0 0 0 2px rgba(245, 158, 11, 0.12)' }
        : undefined
  );
  const dashboardTodayLabel = new Date().toLocaleDateString('uk-UA', { day: '2-digit', month: '2-digit' });
  const openManagerOrder = (orderId: string) => {
    setSelectedOrderId(orderId);
    setPage('orders');
  };
  const openManagerOrderWithFocus = (order: ServiceOrder, target: DashboardFocusTarget) => {
    setDashboardFocus({ orderId: order.id, target });
    setSelectedOrderId(order.id);
    setPage('orders');
  };
  const callClient = (order: ServiceOrder) => {
    recordManagerCall(order);
    if (typeof window !== 'undefined') window.open(`tel:${order.phone}`, '_self');
  };
  const remindEngineer = (order: ServiceOrder) => {
    openManagerOrderWithFocus(order, 'issue');
    if (typeof window !== 'undefined') window.alert(`Нагадування: ${order.engineer} · ${order.id} · у роботі ${engineerElapsedLabel(order.takenInRepairAt ?? order.statusChangedAt)}.`);
  };
  const sendDashboardSms = (order: ServiceOrder, kind: 'wait' | 'ready' | 'unpaid' | 'overdue', days = 0) => {
    if (kind === 'wait') {
      sendClientNotification(order, 'Діагностика завершена', 'SMS', true);
      return;
    }
    if (kind === 'ready') {
      sendClientNotification(order, 'Готово до видачі', 'SMS', true);
      return;
    }
    if (kind === 'unpaid') {
      sendClientNotification(order, 'Очікує оплату', 'SMS', true);
      return;
    }
    sendClientNotification(order, 'Замовлення затримується', 'SMS', true, renderNotificationText(order, 'Замовлення затримується', days));
  };
  const doDashboardNextStep = (order: ServiceOrder, kind: 'wait' | 'ready' | 'unpaid' | 'overdue', days = 0) => {
    if (kind === 'wait') {
      sendDashboardSms(order, kind, days);
      return;
    }
    if (kind === 'ready') {
      openManagerOrderWithFocus(order, 'issue');
      return;
    }
    if (kind === 'unpaid') {
      openManagerOrderWithFocus(order, 'payment');
      return;
    }
    sendDashboardSms(order, kind, days);
  };
  const managerColumnStyle = (background: string, border: string) => ({
    background,
    border: `1px solid ${border}`,
    borderRadius: '20px',
    padding: '18px',
    display: 'flex',
    flexDirection: 'column' as const,
    gap: '14px',
    minHeight: '100%',
  });
  const managerMiniCardStyle = (border: string) => ({
    background: '#fff',
    border: `1px solid ${border}`,
    borderRadius: '16px',
    padding: '14px',
    display: 'flex',
    flexDirection: 'column' as const,
    gap: '10px',
  });
  const dashboardTitle = isEngineer
    ? `Робочий стіл інженера: ${activeUser.name}`
    : isAccountant
      ? 'Робочий стіл бухгалтера'
      : isStock
        ? 'Робочий стіл складу'
      : isManager
        ? 'Робочий стіл менеджера'
        : 'Робочий стіл керівника';

  if (isOwner) {
    const todayDate = parseDateTime(today);
    const todayKey = todayDate ? `${todayDate.getFullYear()}-${todayDate.getMonth()}` : '';
    const isInCurrentMonth = (dateText: string) => {
      const parsed = parseDateTime(dateText);
      if (!parsed || !todayKey) return false;
      return `${parsed.getFullYear()}-${parsed.getMonth()}` === todayKey;
    };
    const paidOrderPayments = orders.flatMap((order) => order.payments
      .filter((payment) => paymentCountsAsApplied(payment) && payment.amount > 0)
      .map((payment) => ({ orderId: order.id, amount: payment.amount, method: payment.method, date: payment.date })));
    const paidSalePayments = sales.flatMap((sale) => sale.payments
      .filter((payment) => paymentCountsAsApplied(payment) && payment.amount > 0)
      .map((payment) => ({ orderId: sale.id, amount: payment.amount, method: payment.method, date: payment.date })));
    const allPaidRows = [...paidOrderPayments, ...paidSalePayments]
      .sort((a, b) => (parseDateTime(b.date)?.getTime() ?? 0) - (parseDateTime(a.date)?.getTime() ?? 0));
    const revenueToday = allPaidRows.filter((payment) => extractDayKey(payment.date) === extractDayKey(today)).reduce((sum, payment) => sum + payment.amount, 0);
    const revenueMonth = allPaidRows.filter((payment) => isInCurrentMonth(payment.date)).reduce((sum, payment) => sum + payment.amount, 0);
    const ordersToday = orders.filter((order) => extractDayKey(order.intakeDate) === extractDayKey(today)).length;
    const activeNow = orders.filter((order) => !['Видано', 'Закрито', 'Скасовано', 'Не підлягає ремонту'].includes(order.status)).length;
    const readyNow = orders.filter((order) => simpleRepairStatus(order.status) === 'Готово').length;
    const averageReceipt = allPaidRows.length > 0 ? Math.round(allPaidRows.reduce((sum, payment) => sum + payment.amount, 0) / allPaidRows.length) : 0;
    const stuckBusinessOrders = orders
      .filter((order) => !['Видано', 'Закрито', 'Скасовано'].includes(order.status) && daysSince(order.statusChangedAt) >= 3)
      .map((order) => ({ order, days: daysSince(order.statusChangedAt) }))
      .sort((a, b) => b.days - a.days)
      .slice(0, 8);
    const totalOrderAmount = orders.reduce((sum, order) => sum + orderTotals(order).total, 0);
    const totalPaidAmount = orders.reduce((sum, order) => sum + orderTotals(order).paid, 0) + sales.reduce((sum, sale) => sum + saleTotals(sale).paid, 0);
    const totalDebtAmount = orders.reduce((sum, order) => sum + orderTotals(order).debt, 0) + sales.reduce((sum, sale) => sum + saleTotals(sale).debt, 0);
    const totalProfitAmount = (analytics.salesProfit ?? 0) + (analytics.serviceProfit ?? 0);
    const directorKpis = [
      { label: 'Замовлення', value: String(orders.length), page: 'orders' as Page },
      { label: 'Виручка', value: money(totalOrderAmount), page: 'finance' as Page },
      { label: 'Оплачено', value: money(totalPaidAmount), page: 'cash' as Page },
      { label: 'Борг', value: money(totalDebtAmount), page: 'problem-clients' as Page },
      { label: 'Прибуток', value: money(totalProfitAmount), page: 'finance' as Page },
    ];
    const debtorRows = customerList
      .map((client) => {
        const clientOrders = orders.filter((order) => isSameClientOrder(client, order) && clientOrderDebt(order) > 0);
        const amount = clientOrders.reduce((sum, order) => sum + clientOrderDebt(order), 0);
        const maxDays = clientOrders.reduce((max, order) => Math.max(max, daysSince(order.debtSince ?? order.intakeDate)), 0);
        return { client, amount, maxDays };
      })
      .filter((row) => row.amount > 0)
      .sort((a, b) => b.maxDays - a.maxDays || b.amount - a.amount)
      .slice(0, 8);
    const unpaidDirectorOrders = orders
      .filter((order) => orderTotals(order).debt > 0)
      .sort((a, b) => daysSince(b.debtSince ?? b.intakeDate) - daysSince(a.debtSince ?? a.intakeDate))
      .slice(0, 8);
    const noEngineerDirectorOrders = orders.filter((order) => !order.engineer).slice(0, 8);
    const belowMinProducts = products
      .filter((product) => available(product) < Math.max(product.min, 1))
      .sort((a, b) => (available(a) - a.min) - (available(b) - b.min))
      .slice(0, 8);
    const incomingToday = simplePayments
      .filter((payment) => payment.status === 'подтвержден' && payment.direction !== 'outgoing' && extractDayKey(payment.date) === extractDayKey(today))
      .reduce((sum, payment) => sum + payment.amount, 0);
    const outgoingToday = purchases
      .filter((purchase) => purchase.purchasedAt && extractDayKey(purchase.purchasedAt) === extractDayKey(today))
      .reduce((sum, purchase) => sum + purchase.items.reduce((inner, item) => inner + ((item.received || item.qty) * item.price), 0), 0);
    const balanceToday = incomingToday - outgoingToday;
    const bankRows = companyBankAccounts.map((account) => ({
      account,
      incoming: simplePayments
        .filter((payment) => payment.status === 'подтвержден' && payment.direction !== 'outgoing' && payment.bankAccount === account.iban)
        .reduce((sum, payment) => sum + payment.amount, 0),
    }));
    const topServiceRows = Array.from(
      orders.reduce((map, order) => {
        order.works.forEach((work) => {
          const key = work.serviceType || work.name;
          const current = map.get(key) ?? { label: key, amount: 0 };
          current.amount += work.price * (work.qty ?? 1);
          map.set(key, current);
        });
        return map;
      }, new Map<string, { label: string; amount: number }>()),
    ).map(([, value]) => value).sort((a, b) => b.amount - a.amount).slice(0, 5);

    return (
      <div className="page-grid executive-dashboard">
        <PageTitle eyebrow="Головна" title="Контроль бізнесу" text="Гроші, проблеми і точки уваги." />

        <section className="executive-summary-strip">
          {directorKpis.map((item) => (
            <button key={item.label} type="button" className="executive-summary-item" onClick={() => setPage(item.page)}>
              <span>{item.label}</span>
              <strong>{item.value}</strong>
            </button>
          ))}
        </section>

        <section className="panel executive-panel">
          <div className="panel-heading">
            <h2>Проблеми</h2>
            <span>куди дивитись</span>
          </div>
          <div className="executive-grid">
            <div className="executive-block">
              <div className="executive-block-head">
                <strong>Борги клієнтів</strong>
                <button type="button" className="executive-link-button" onClick={() => setPage('problem-clients')}>Відкрити</button>
              </div>
              <div className="executive-list executive-alert-list">
                {debtorRows.map((row) => (
                  <button key={row.client.phone} type="button" className="executive-order-row executive-order-alert executive-row-button" onClick={() => setPage('problem-clients')}>
                    <div>
                      <strong>{row.client.name}</strong>
                      <span>{row.maxDays} дн.</span>
                    </div>
                    <div>
                      <strong>{money(row.amount)}</strong>
                    </div>
                  </button>
                ))}
                {debtorRows.length === 0 && <div className="empty-state">Боргів немає.</div>}
              </div>
            </div>

            <div className="executive-block">
              <div className="executive-block-head">
                <strong>Без оплати</strong>
                <button type="button" className="executive-link-button" onClick={() => setPage('orders')}>Відкрити</button>
              </div>
              <div className="executive-list">
                {unpaidDirectorOrders.map((order) => (
                  <button key={order.id} type="button" className="executive-order-row executive-row-button" onClick={() => { setSelectedOrderId(order.id); setPage('orders'); }}>
                    <div>
                      <strong>{order.id}</strong>
                      <span>{order.client}</span>
                    </div>
                    <div>
                      <strong>{money(orderTotals(order).debt)}</strong>
                    </div>
                  </button>
                ))}
                {unpaidDirectorOrders.length === 0 && <div className="empty-state">Неоплачених немає.</div>}
              </div>
            </div>

            <div className="executive-block">
              <div className="executive-block-head">
                <strong>Без інженера</strong>
                <button type="button" className="executive-link-button" onClick={() => setPage('orders')}>Відкрити</button>
              </div>
              <div className="executive-list">
                {noEngineerDirectorOrders.map((order) => (
                  <button key={order.id} type="button" className="executive-order-row executive-row-button" onClick={() => { setSelectedOrderId(order.id); setPage('orders'); }}>
                    <div>
                      <strong>{order.id}</strong>
                      <span>{order.device}</span>
                    </div>
                    <div>
                      <span>{order.client}</span>
                    </div>
                  </button>
                ))}
                {noEngineerDirectorOrders.length === 0 && <div className="empty-state">Усі з інженером.</div>}
              </div>
            </div>

            <div className="executive-block">
              <div className="executive-block-head">
                <strong>Нижче мінімуму</strong>
                <button type="button" className="executive-link-button" onClick={() => setPage('parts')}>Відкрити</button>
              </div>
              <div className="executive-list">
                {belowMinProducts.map((product) => (
                  <button key={product.id} type="button" className="executive-order-row executive-row-button" onClick={() => setPage('parts')}>
                    <div>
                      <strong>{product.sku}</strong>
                      <span>{product.name}</span>
                    </div>
                    <div>
                      <strong>{available(product)}/{product.min}</strong>
                    </div>
                  </button>
                ))}
                {belowMinProducts.length === 0 && <div className="empty-state">Склад у нормі.</div>}
              </div>
            </div>
          </div>
        </section>

        <section className="executive-grid">
          <section className="panel executive-panel">
            <div className="panel-heading">
              <h2>Гроші за день</h2>
              <span>{extractDayKey(today)}</span>
            </div>
            <div className="executive-money-grid">
              <button type="button" className="executive-summary-item" onClick={() => setPage('cash')}>
                <span>Прихід</span>
                <strong>{money(incomingToday)}</strong>
              </button>
              <button type="button" className="executive-summary-item" onClick={() => setPage('purchases')}>
                <span>Витрата</span>
                <strong>{money(outgoingToday)}</strong>
              </button>
              <button type="button" className="executive-summary-item" onClick={() => setPage('finance')}>
                <span>Баланс</span>
                <strong>{money(balanceToday)}</strong>
              </button>
            </div>
            <div className="executive-list">
              {bankRows.map((row) => (
                <button key={row.account.id} type="button" className="executive-list-row executive-row-button" onClick={() => setPage('bank-import')}>
                  <span>{row.account.bankName}</span>
                  <strong>{money(row.incoming)}</strong>
                </button>
              ))}
            </div>
          </section>

          <section className="panel executive-panel">
            <div className="panel-heading">
              <h2>Мінімум аналітики</h2>
              <span>де заробіток і де затримка</span>
            </div>
            <div className="executive-block">
              <div className="executive-block-head">
                <strong>Топ прибуткових послуг</strong>
                <button type="button" className="executive-link-button" onClick={() => setPage('finance')}>Відкрити</button>
              </div>
              <div className="executive-list">
                {topServiceRows.map((row) => (
                  <button key={row.label} type="button" className="executive-list-row executive-row-button" onClick={() => setPage('finance')}>
                    <span>{row.label}</span>
                    <strong>{money(row.amount)}</strong>
                  </button>
                ))}
                {topServiceRows.length === 0 && <div className="empty-state">Даних ще немає.</div>}
              </div>
            </div>
            <div className="executive-block">
              <div className="executive-block-head">
                <strong>Зависші замовлення</strong>
                <button type="button" className="executive-link-button" onClick={() => setPage('orders')}>Відкрити</button>
              </div>
              <div className="executive-list executive-alert-list">
                {stuckBusinessOrders.map(({ order, days }) => (
                  <button key={order.id} type="button" className="executive-order-row executive-order-alert executive-row-button" onClick={() => { setSelectedOrderId(order.id); setPage('orders'); }}>
                    <div>
                      <strong>{order.id}</strong>
                      <span>{order.client}</span>
                    </div>
                    <div>
                      <span>{managerOrderStatusLabel(order.status)}</span>
                      <strong>{days} дн.</strong>
                    </div>
                  </button>
                ))}
                {stuckBusinessOrders.length === 0 && <div className="empty-state">Зависших немає.</div>}
              </div>
            </div>
          </section>
        </section>
      </div>
    );
  }

  return (
    <div className="page-grid">
      <PageTitle eyebrow="Головна" title={dashboardTitle} text="Перший екран для щоденної роботи: швидкі дії, поточні задачі, сигнали і гроші без зайвих переходів." />
      {(isManager || isOwner) && (
        <QuickOrderIntake
          phone={quickPhone}
          clientName={quickClientName}
          device={quickDevice}
          serial={quickSerial}
          problem={quickProblem}
          appearance={quickAppearance}
          estimatedAmount={quickEstimatedAmount}
          engineerId={quickEngineerId}
          contractId={quickContractId}
          locationCode={quickLocationCode}
          comment={quickComment}
          managerName={activeUser.name}
          users={users}
          contracts={contracts}
          customerList={customerList}
          warehouseLocations={warehouseLocations}
          onPhoneChange={setQuickPhone}
          onClientNameChange={setQuickClientName}
          onDeviceChange={setQuickDevice}
          onSerialChange={setQuickSerial}
          onProblemChange={setQuickProblem}
          onAppearanceChange={setQuickAppearance}
          onEstimatedAmountChange={setQuickEstimatedAmount}
          onEngineerChange={setQuickEngineerId}
          onContractChange={setQuickContractId}
          onLocationChange={setQuickLocationCode}
          onCommentChange={setQuickComment}
          onSubmit={createQuickOrder}
        />
      )}

      {isManager && (
        <AutoControlPanel
          orders={orders}
          sales={sales}
          users={users}
          title="Автоконтроль"
        />
      )}
     
      {(isManager || isOwner) && (
        <section className="panel">
          <div className="panel-heading">
            <h2>Контроль помилок</h2>
            <span>{criticalProblems.length} критично · автооновлення 5-10 хв</span>
          </div>
          <div className="table problem-table">
            <div className="table-row table-head"><span>Рівень</span><span>Заказ</span><span>Проблема</span><span>Днів</span><span>Відповідальний</span><span>Дія</span></div>
            {roleServiceProblems.slice(0, 5).map((problem) => (
              <div className="table-row" key={`dash-${problem.id}`}>
                <span>{problem.level}</span>
                <span>{problem.orderId}<small>{problem.client}</small></span>
                <span>{problem.type}<small>{problem.description}</small></span>
                <span>{problem.days}</span>
                <span>{problem.responsible}</span>
                <span><button type="button" onClick={() => { setSelectedOrderId(problem.orderId); setPage('orders'); }}>Відкрити заказ</button></span>
              </div>
            ))}
          </div>
          {roleServiceProblems.length === 0 && <div className="empty-state">Критичних проблем зараз немає.</div>}
        </section>
      )}
      {isManager && (
        <>
          <section className="stats-grid">
            <Metric icon={<History />} label={`Сьогодні: ${dashboardTodayLabel}`} value={String(managerOrders.length)} hint="усього замовлень менеджера" />
            <Metric icon={<ClipboardList />} label="В роботі" value={String(managerOrders.filter((order) => !['Готовий до видачі', 'Очікує оплати', 'Очікує клієнта', 'Видано', 'Закрито', 'Скасовано'].includes(order.status)).length)} hint="активні замовлення" />
            <Metric icon={<PackageCheck />} label="Готово" value={String(managerReadyNotIssuedOrders.length)} hint="чекають клієнта" />
            <Metric icon={<Banknote />} label="Гроші чекають" value={String(managerUnpaidOrders.length)} hint="не оплачено" />
          </section>
          <section className="stats-grid">
            <Metric icon={<Banknote />} label="Сьогодні" value={money(managerMoneyToday)} hint="прийнято оплат за день" />
            <Metric icon={<History />} label="Борги" value={money(managerMoneyDebt)} hint="усього висить по замовленнях" />
            <Metric icon={<ClipboardList />} label="Частково" value={money(managerMoneyPartial)} hint="залишок після часткових оплат" />
            <Metric icon={<X />} label="Прострочено" value={money(managerMoneyOverdue)} hint="ризикові гроші та мовчазні клієнти" />
          </section>
          <section style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(250px, 1fr))', gap: '16px' }}>
            <div style={managerColumnStyle('#fff7d6', '#facc15')}>
              <div className="panel-heading"><h2>Ждут ответа</h2><span>{managerWaitResponseOrders.length}</span></div>
              {managerWaitResponseOrders.slice(0, 6).map(({ order, approval }) => (
                <article key={`wait-${order.id}`} style={{ ...managerMiniCardStyle('#facc15'), ...managerUrgencyStyle(daysSince(approval.sentAt)) }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', gap: '12px', alignItems: 'flex-start' }}>
                    <strong style={{ fontSize: '18px' }}>{order.id}</strong>
                    <span style={{ fontSize: '18px', fontWeight: 800, color: '#a16207' }}>{money(approval.totalAmount)}</span>
                  </div>
                  <div style={{ fontWeight: 600 }}>{order.device}</div>
                  <div style={{ fontSize: '20px', fontWeight: 800 }}>{engineerElapsedLabel(approval.sentAt)}</div>
                  <div className="action-row">
                    <button type="button" onClick={() => doDashboardNextStep(order, 'wait')}>Наступний крок</button>
                    <button type="button" onClick={() => sendDashboardSms(order, 'wait')}>SMS</button>
                    <button type="button" onClick={() => openManagerOrderWithFocus(order, 'approval')}>Відкрити</button>
                    <button type="button" onClick={() => callClient(order)}>Подзвонити</button>
                  </div>
                </article>
              ))}
              {managerWaitResponseOrders.length === 0 && <div className="empty-state">Порожньо.</div>}
            </div>
            <div style={managerColumnStyle('#eaffef', '#22c55e')}>
              <div className="panel-heading"><h2>Готово, але не забрали</h2><span>{managerReadyNotIssuedOrders.length}</span></div>
              {managerReadyNotIssuedOrders.slice(0, 6).map(({ order, ageDays }) => (
                <article key={`ready-${order.id}`} style={{ ...managerMiniCardStyle('#22c55e'), ...managerUrgencyStyle(ageDays) }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', gap: '12px', alignItems: 'flex-start' }}>
                    <strong style={{ fontSize: '18px' }}>{order.id}</strong>
                    <span style={{ fontSize: '18px', fontWeight: 800, color: '#15803d' }}>{money(orderTotals(order).total)}</span>
                  </div>
                  <div style={{ fontWeight: 600 }}>{order.device}</div>
                  <div>Комірка: <strong>{order.locationCode ?? 'не призначено'}</strong></div>
                  <div style={{ fontSize: '20px', fontWeight: 800 }}>{ageDays} дн.</div>
                  <div className="action-row">
                    <button type="button" onClick={() => doDashboardNextStep(order, 'ready', ageDays)}>Наступний крок</button>
                    <button type="button" onClick={() => sendDashboardSms(order, 'ready', ageDays)}>SMS</button>
                    <button type="button" onClick={() => openManagerOrderWithFocus(order, 'issue')}>Відкрити</button>
                    <button type="button" onClick={() => callClient(order)}>Подзвонити</button>
                  </div>
                </article>
              ))}
              {managerReadyNotIssuedOrders.length === 0 && <div className="empty-state">Порожньо.</div>}
            </div>
            <div style={managerColumnStyle('#eaf4ff', '#38bdf8')}>
              <div className="panel-heading"><h2>Не оплачено</h2><span>{managerUnpaidOrders.length}</span></div>
              {managerUnpaidOrders.slice(0, 6).map(({ order, ageDays }) => (
                <article key={`unpaid-${order.id}`} style={{ ...managerMiniCardStyle('#38bdf8'), ...managerUrgencyStyle(ageDays) }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', gap: '12px', alignItems: 'flex-start' }}>
                    <strong style={{ fontSize: '18px' }}>{order.id}</strong>
                  </div>
                  <div style={{ fontSize: '24px', fontWeight: 900, color: '#0369a1' }}>{money(orderTotals(order).debt)}</div>
                  <div style={{ fontWeight: 600 }}>До сплати</div>
                  <div className="action-row">
                    <button type="button" onClick={() => doDashboardNextStep(order, 'unpaid', ageDays)}>Наступний крок</button>
                    <button type="button" onClick={() => sendDashboardSms(order, 'unpaid', ageDays)}>SMS</button>
                    <button type="button" onClick={() => openManagerOrderWithFocus(order, 'payment')}>Відкрити</button>
                    <button type="button" onClick={() => callClient(order)}>Подзвонити</button>
                  </div>
                </article>
              ))}
              {managerUnpaidOrders.length === 0 && <div className="empty-state">Порожньо.</div>}
            </div>
            <div style={managerColumnStyle('#ffe8e8', '#ef4444')}>
              <div className="panel-heading"><h2>Прострочено</h2><span>{managerOverdueOrders.length}</span></div>
              {managerOverdueOrders.slice(0, 6).map(({ order, ageDays }) => (
                <article key={`overdue-${order.id}`} style={{ ...managerMiniCardStyle('#ef4444'), background: '#fff5f5', ...managerUrgencyStyle(ageDays) }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', gap: '12px', alignItems: 'flex-start' }}>
                    <strong style={{ fontSize: '18px' }}>{order.id}</strong>
                    <span style={{ fontSize: '18px', fontWeight: 800, color: '#b91c1c' }}>{money(orderTotals(order).debt || orderTotals(order).total)}</span>
                  </div>
                  <div style={{ fontSize: '22px', fontWeight: 900, color: '#b91c1c' }}>{ageDays} дн.</div>
                  <div style={{ fontWeight: 600 }}>{order.device}</div>
                  <div className="action-row">
                    <button type="button" onClick={() => doDashboardNextStep(order, 'overdue', ageDays)}>Наступний крок</button>
                    <button type="button" onClick={() => sendDashboardSms(order, 'overdue', ageDays)}>SMS</button>
                    <button type="button" onClick={() => openManagerOrderWithFocus(order, 'issue')}>Відкрити</button>
                    <button type="button" onClick={() => callClient(order)}>Подзвонити</button>
                  </div>
                </article>
              ))}
              {managerOverdueOrders.length === 0 && <div className="empty-state">Порожньо.</div>}
            </div>
          </section>
          <section style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(250px, 1fr))', gap: '16px', marginTop: '16px' }}>
            <div style={managerColumnStyle('#fff7eb', '#fb923c')}>
              <div className="panel-heading"><h2>Деньги ждут</h2><span>{managerMoneyWaitOrders.length}</span></div>
              {managerMoneyWaitOrders.slice(0, 6).map(({ order, debt, ageDays }) => (
                <article key={`money-wait-${order.id}`} style={{ ...managerMiniCardStyle('#fb923c'), ...managerUrgencyStyle(ageDays) }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', gap: '12px' }}>
                    <strong style={{ fontSize: '18px' }}>{order.id}</strong>
                    <span style={{ fontSize: '22px', fontWeight: 900, color: '#c2410c' }}>{money(debt)}</span>
                  </div>
                  <div style={{ fontWeight: 600 }}>{ageDays} дн. без оплати</div>
                  <div className="action-row">
                    <button type="button" onClick={() => sendDashboardSms(order, 'unpaid', ageDays)}>SMS</button>
                    <button type="button" onClick={() => callClient(order)}>Подзвонити</button>
                    <button type="button" onClick={() => doDashboardNextStep(order, 'unpaid', ageDays)}>Видати</button>
                  </div>
                </article>
              ))}
              {managerMoneyWaitOrders.length === 0 && <div className="empty-state">Порожньо.</div>}
            </div>
            <div style={managerColumnStyle('#eef6ff', '#60a5fa')}>
              <div className="panel-heading"><h2>Частично оплачено</h2><span>{managerPartialPaidOrders.length}</span></div>
              {managerPartialPaidOrders.slice(0, 6).map(({ order, paid, debt, lastPayment, ageDays }) => (
                <article key={`money-partial-${order.id}`} style={{ ...managerMiniCardStyle('#60a5fa'), ...managerUrgencyStyle(ageDays) }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', gap: '12px' }}>
                    <strong style={{ fontSize: '18px' }}>{order.id}</strong>
                    <span style={{ fontSize: '20px', fontWeight: 900, color: '#2563eb' }}>{money(debt)}</span>
                  </div>
                  <div>Оплачено: {money(paid)}</div>
                  <div>Залишилось: {money(debt)}</div>
                  <div style={{ fontWeight: 600 }}>{lastPayment ? `Платили ${lastPayment.date}` : `${ageDays} дн.`}</div>
                  <div className="action-row">
                    <button type="button" onClick={() => openManagerOrderWithFocus(order, 'payment')}>Добити оплату</button>
                  </div>
                </article>
              ))}
              {managerPartialPaidOrders.length === 0 && <div className="empty-state">Порожньо.</div>}
            </div>
            <div style={managerColumnStyle('#ecfdf5', '#34d399')}>
              <div className="panel-heading"><h2>Оплачено сьогодні</h2><span>{money(managerPaidTodayTotal)}</span></div>
              {managerPaidTodayRows.slice(0, 6).map(({ order, payment }) => (
                <article key={`money-today-${payment.id}`} style={managerMiniCardStyle('#34d399')}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', gap: '12px' }}>
                    <strong style={{ fontSize: '18px' }}>{order.id}</strong>
                    <span style={{ fontSize: '20px', fontWeight: 900, color: '#047857' }}>{money(payment.amount)}</span>
                  </div>
                  <div>{payment.method} · {payment.date}</div>
                  <div>Прийняв: {payment.acceptedBy}</div>
                </article>
              ))}
              {managerPaidTodayRows.length === 0 && <div className="empty-state">Сьогодні оплат ще не було.</div>}
            </div>
            <div style={managerColumnStyle('#fef2f2', '#ef4444')}>
              <div className="panel-heading"><h2>Проблемные деньги</h2><span>{managerProblemMoneyOrders.length}</span></div>
              {managerProblemMoneyOrders.slice(0, 6).map(({ order, debt, ageDays }) => (
                <article key={`money-problem-${order.id}`} style={{ ...managerMiniCardStyle('#ef4444'), background: '#fff5f5', ...managerUrgencyStyle(ageDays) }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', gap: '12px' }}>
                    <strong style={{ fontSize: '18px' }}>{order.id}</strong>
                    <span style={{ fontSize: '20px', fontWeight: 900, color: '#b91c1c' }}>{money(debt)}</span>
                  </div>
                  <div style={{ fontWeight: 700, color: '#b91c1c' }}>{ageDays} дн. ризику</div>
                  <div className="action-row">
                    <button type="button" onClick={() => sendDashboardSms(order, orderTotals(order).paid > 0 ? 'overdue' : 'unpaid', ageDays)}>SMS</button>
                    <button type="button" onClick={() => callClient(order)}>Подзвонити</button>
                    <button type="button" onClick={() => openManagerOrderWithFocus(order, orderTotals(order).paid > 0 ? 'payment' : 'approval')}>Відкрити</button>
                  </div>
                </article>
              ))}
              {managerProblemMoneyOrders.length === 0 && <div className="empty-state">Порожньо.</div>}
            </div>
          </section>
          {managerCashierRows.length > 0 && (
            <section className="panel" style={{ marginTop: '16px' }}>
              <div className="panel-heading"><h2>Хто прийняв гроші сьогодні</h2><span>{managerCashierRows.length} співробітників</span></div>
              <div className="task-list">
                {managerCashierRows.map(([name, amount]) => (
                  <Task key={`cashier-${name}`} icon={<Users />} title={name} text={`${money(amount)} прийнято сьогодні`} />
                ))}
              </div>
            </section>
          )}
        </>
      )}
      {isOwner && (
        <>
          <section className="stats-grid">
            <Metric icon={<X />} label="Просрочено" value={String(ownerOverdueOrders.length)} hint="замовлення старші за контрольний строк" />
            <Metric icon={<Banknote />} label="Долги" value={money(ownerLeakDebtTotal)} hint="усі неоплачені замовлення" />
            <Metric icon={<Wrench />} label="У інженерів" value={String(ownerEngineerActiveCount)} hint="зараз у роботі" />
          </section>
          <section className="panel">
            <div className="panel-heading"><h2>Де губимо гроші</h2><span>черга проблем, а не просто список</span></div>
            <div className="action-row">
              <button type="button" onClick={() => setShowOwnerLeakQueue((current) => !current)}>{showOwnerLeakQueue ? 'Сховати чергу проблем' : 'Розвʼязати все'}</button>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))', gap: '16px', marginTop: '16px' }}>
              <div style={managerColumnStyle('#fff0f0', '#ef4444')}>
                <div className="panel-heading"><h2>Прострочки</h2><span>{ownerOverdueOrders.length}</span></div>
                {ownerOverdueOrders.slice(0, 5).map(({ order, ageDays }) => (
                  <article key={`owner-overdue-${order.id}`} style={{ ...managerMiniCardStyle('#ef4444'), background: '#fff5f5', ...managerUrgencyStyle(ageDays) }}>
                    <strong style={{ fontSize: '18px' }}>{order.id}</strong>
                    <div style={{ fontSize: '20px', fontWeight: 900, color: '#b91c1c' }}>{ageDays} дн.</div>
                    <div>{money(orderTotals(order).debt || orderTotals(order).total)}</div>
                    <div>{order.status}</div>
                    <div className="action-row">
                      <button type="button" onClick={() => callClient(order)}>Подзвонити</button>
                      <button type="button" onClick={() => sendDashboardSms(order, 'overdue', ageDays)}>SMS</button>
                      <button type="button" onClick={() => openManagerOrder(order.id)}>Відкрити</button>
                    </div>
                  </article>
                ))}
                {ownerOverdueOrders.length === 0 && <div className="empty-state">Порожньо.</div>}
              </div>
              <div style={managerColumnStyle('#fff9db', '#facc15')}>
                <div className="panel-heading"><h2>Зависло у інженера</h2><span>{ownerEngineerStuckOrders.length}</span></div>
                {ownerEngineerStuckOrders.slice(0, 5).map(({ order }) => (
                  <article key={`owner-engineer-${order.id}`} style={{ ...managerMiniCardStyle('#facc15'), ...managerUrgencyStyle(Math.floor(hoursSince(order.takenInRepairAt ?? order.statusChangedAt) / 24)) }}>
                    <strong style={{ fontSize: '18px' }}>{order.id}</strong>
                    <div>{order.engineer}</div>
                    <div style={{ fontWeight: 700 }}>{engineerElapsedLabel(order.takenInRepairAt ?? order.statusChangedAt)}</div>
                    <div>{order.device}</div>
                    <div>{order.issue}</div>
                    <div className="action-row">
                      <button type="button" onClick={() => remindEngineer(order)}>Нагадати інженеру</button>
                      <button type="button" onClick={() => openManagerOrder(order.id)}>Відкрити заказ</button>
                    </div>
                  </article>
                ))}
                {ownerEngineerStuckOrders.length === 0 && <div className="empty-state">Порожньо.</div>}
              </div>
              <div style={managerColumnStyle('#eef6ff', '#38bdf8')}>
                <div className="panel-heading"><h2>Неоплачені</h2><span>{ownerUnpaidReadyOrders.length}</span></div>
                {ownerUnpaidReadyOrders.slice(0, 5).map(({ order, debt, ageDays }) => (
                  <article key={`owner-unpaid-${order.id}`} style={{ ...managerMiniCardStyle('#38bdf8'), ...managerUrgencyStyle(ageDays) }}>
                    <strong style={{ fontSize: '18px' }}>{order.id}</strong>
                    <div style={{ fontSize: '20px', fontWeight: 900, color: '#0369a1' }}>{money(debt)}</div>
                    <div>{ageDays} дн. без оплати</div>
                    <div>Комірка: {order.locationCode ?? 'не задано'}</div>
                    <div className="action-row">
                      <button type="button" onClick={() => sendDashboardSms(order, 'unpaid', ageDays)}>SMS оплатити</button>
                      <button type="button" onClick={() => openManagerOrderWithFocus(order, 'payment')}>Відкрити касу</button>
                    </div>
                  </article>
                ))}
                {ownerUnpaidReadyOrders.length === 0 && <div className="empty-state">Порожньо.</div>}
              </div>
              <div style={managerColumnStyle('#ecfdf5', '#22c55e')}>
                <div className="panel-heading"><h2>Готово, але не забрали</h2><span>{ownerReadyNotPickedOrders.length}</span></div>
                {ownerReadyNotPickedOrders.slice(0, 5).map(({ order, ageDays }) => (
                  <article key={`owner-ready-${order.id}`} style={{ ...managerMiniCardStyle('#22c55e'), ...managerUrgencyStyle(ageDays) }}>
                    <strong style={{ fontSize: '18px' }}>{order.id}</strong>
                    <div>{ageDays} дн. лежить</div>
                    <div>{money(orderTotals(order).total)}</div>
                    <div>Комірка: {order.locationCode ?? 'не задано'}</div>
                    <div className="action-row">
                      <button type="button" onClick={() => sendDashboardSms(order, 'ready', ageDays)}>SMS забрати</button>
                      <button type="button" onClick={() => callClient(order)}>Подзвонити</button>
                    </div>
                  </article>
                ))}
                {ownerReadyNotPickedOrders.length === 0 && <div className="empty-state">Порожньо.</div>}
              </div>
            </div>
          </section>
          {showOwnerLeakQueue && (
            <section className="panel">
              <div className="panel-heading"><h2>Черга проблем</h2><span>{ownerLeakQueue.length} замовлень підряд</span></div>
              <div className="task-list">
                {ownerLeakQueue.map((item) => (
                  <Task
                    key={`owner-queue-${item.type}-${item.order.id}`}
                    icon={<Bell />}
                    title={`${item.type} · ${item.order.id} · ${money(item.amount)}`}
                    text={`${item.order.client} · ${item.order.device} · ${item.ageDays} дн. · статус ${item.order.status} · комірка ${item.order.locationCode ?? 'не задано'}`}
                  />
                ))}
              </div>
            </section>
          )}
          <section className="stats-grid">
            <Metric icon={<Banknote />} label="Виручка сьогодні" value={money(salesToday + serviceRevenue)} hint="оплати ремонтів + продажі" />
            <Metric icon={<Banknote />} label="Виручка за місяць" value={money(monthRevenue)} hint="демо-період у CRM" />
            <Metric icon={<CheckCircle2 />} label="Прибуток" value={money(monthProfit)} hint="тільки фінальні ремонти + продажі" />
            <Metric icon={<History />} label="Долги" value={money(debt)} hint="клієнти і продажі" />
          </section>
          <section className="content-split">
            <div className="panel">
              <div className="panel-heading"><h2>Проблеми</h2><span>для рішення зараз</span></div>
              <div className="task-list">
                <Task icon={<History />} title="Заказы без руху" text={`${stuckOrders.length} не змінювали статус більше 2 днів.`} />
                <Task icon={<Banknote />} title="Немає оплати" text={`${noPaymentOrders.length} заказів із боргом.`} />
                <Task icon={<ClipboardList />} title="Немає актів" text={`${noActOrders.length} актів не повернено.`} />
                <Task icon={<X />} title="Помилки ПН" text={`${overdueTaxInvoices.length} податкових накладних потребують уваги.`} />
                <Task icon={<PackagePlus />} title="Немає запчастей" text={`${missingParts.length} потреб треба закупити.`} />
              </div>
            </div>
            <div className="panel">
              <div className="panel-heading"><h2>Операції</h2><span>стан сервісу</span></div>
              <div className="details-grid">
                <Info label="Активні заказы" value={String(activeRepairs)} />
                <Info label="В ремонті" value={String(inRepair)} />
                <Info label="Готово" value={String(ready)} />
                <Info label="На видачу" value={String(roleReadyOrders.length)} />
              </div>
            </div>
          </section>
          <section className="content-split">
            <div className="panel">
              <div className="panel-heading"><h2>Сотрудники KPI</h2><span>менеджери, інженери, бухгалтерія</span></div>
              <div className="task-list">
                {managerRows.map((row) => <Task key={`owner-manager-${row.name}`} icon={<Users />} title={`Менеджер · ${row.name}`} text={`Прийнято ${row.orders}, видано ${row.closed}, оплат ${row.paymentsCount}, середній чек ${money(row.avg)}, борги ${money(row.debt)}, актів не повернено ${row.documentsCount}.`} />)}
                {engineerRows.map((row) => <Task key={`owner-engineer-${row.name}`} icon={<Wrench />} title={`Інженер · ${row.name}`} text={`Заказів ${row.orders}, завершено ${row.finished}, середній час ${row.avgDays} дн., без руху ${row.stuck}, повтори ${row.repeated}, заробіток ${money(row.earning)}, прибуток компанії ${money(row.companyProfit)}.`} />)}
                <Task icon={<ClipboardList />} title="Бухгалтер" text={`Створено ПН ${accountantKpi.readyTax}, помилок ${accountantKpi.overdueTax}, акти без повернення ${accountantKpi.acts}, неоплачені заказы ${accountantKpi.unpaid}.`} />
              </div>
            </div>
            <div className="panel">
              <div className="panel-heading"><h2>Склад</h2><span>залишки і закупки</span></div>
              <div className="task-list">
                <Task icon={<Archive />} title="Остатки" text={`${stockKpi.stock} одиниць доступно на складі.`} />
                <Task icon={<X />} title="Критичні залишки" text={`${stockKpi.critical} позицій нижче мінімуму.`} />
                <Task icon={<Truck />} title="Приходи" text={`${stockKpi.receipts} закупок прибули або частково прибули.`} />
                <Task icon={<History />} title="Затримки поставок" text={`${stockKpi.delays} закупок прострочені за очікуваною датою.`} />
              </div>
            </div>
          </section>
          <section className="panel">
            <div className="panel-heading"><h2>Зарплата сотрудников</h2><span>{payrollDashboardRows.length} сотрудников</span></div>
            <div className="table payroll-table">
              <div className="table-row table-head"><span>Сотрудник</span><span>За сегодня</span><span>За месяц</span><span>Выполнено работ</span></div>
              {payrollDashboardRows.map((row) => (
                <div className="table-row" key={`dash-payroll-${row.rule.employee}`}>
                  <span>{row.rule.employee}</span>
                  <span>{money(row.dayAccrued)}</span>
                  <span>{money(row.monthAccrued)}</span>
                  <span>{row.monthOperations}</span>
                </div>
              ))}
            </div>
          </section>
          <section className="panel">
            <div className="panel-heading"><h2>Уведомлення керівнику</h2><span>проблемні події</span></div>
            <div className="task-list">
              {[...problemOrders.slice(0, 3).map((order) => <Task key={`problem-${order.id}`} icon={<Bell />} title={`${order.id} · ${order.status}`} text={`${order.client}: ${order.issue}`} />),
                ...overdueTaxInvoices.slice(0, 2).map((invoice) => <Task key={`tax-${invoice.id}`} icon={<ClipboardList />} title={`${invoice.number} · помилка ПН`} text={`${invoice.client}: ${money(invoice.amount)} · створено ${invoice.createdAt}`} />),
                ...minStock.slice(0, 2).map((product) => <Task key={`stock-${product.id}`} icon={<Archive />} title={`${product.name} · критичний залишок`} text={`Доступно ${available(product)} шт., мінімум ${product.min}.`} />),
              ]}
            </div>
          </section>
        </>
      )}
      <section className="content-split">
        <div className="panel">
          <div className="panel-heading"><h2>Поточна робота</h2><span>{roleOrders.length} заказів</span></div>
          <div className="task-list">
            {isStock ? (
              <>
                <Task icon={<Truck />} title="Прихід" text={`${purchases.length} закупівель у контролі складу.`} />
                <Task icon={<History />} title="Переміщення" text="Рухи складу відкриваються в один клік." />
                <Task icon={<Archive />} title="Остатки" text={`${minStock.length} позицій нижче мінімального залишку.`} />
              </>
            ) : (
              <>
                <Task icon={<Wrench />} title="Активні ремонти" text={`${roleActiveRepairs.length} у роботі зараз.`} />
                <Task icon={<PackageCheck />} title="Готово до видачі" text={`${roleReadyOrders.length} можна видавати або контролювати клієнта.`} />
                <Task icon={<History />} title="Проблемні заказы" text={`${problemOrders.length} зависли, прострочені або довго чекають запчастину.`} />
              </>
            )}
          </div>
        </div>
        <div className="panel">
          <div className="panel-heading"><h2>Сигнали</h2><span>дії сьогодні</span></div>
          <div className="task-list">
            {isStock ? (
              <>
                <Task icon={<Archive />} title="Критичні залишки" text={`${minStock.length} товарів потребують закупки або контролю.`} />
                <Task icon={<PackagePlus />} title="Потреби сервісу" text={`${requirements.length} потреб із ремонтів очікують забезпечення.`} />
                <Task icon={<Truck />} title="В дорозі" text={`${analytics.ordered} одиниць замовлено або очікується.`} />
              </>
            ) : (
              <>
                {(isManager || isAccountant || isOwner) && <Task icon={<Banknote />} title="Немає оплати" text={`${noPaymentOrders.length} заказів або продажів мають борг.`} />}
                {(isManager || isAccountant || isOwner) && <Task icon={<ClipboardList />} title="Немає актів" text={`${noActOrders.length} актів не повернені або потребують контролю.`} />}
                <Task icon={<History />} title="Завислі заказы" text={`${stuckOrders.length} без руху більше 2 днів.`} />
              </>
            )}
          </div>
        </div>
      </section>
     
      {(isManager || isAccountant || isOwner || isStock) && <section className="content-split">
        <div className="panel">
          <div className="panel-heading"><h2>{isStock ? 'Складські проблеми' : 'Де проблеми'}</h2><span>{isStock ? minStock.length : problemOrders.length} сигналів</span></div>
          <div className="task-list">
            {isStock ? (
              <>
                <Task icon={<Archive />} title="Мінімальні залишки" text={`${minStock.length} позицій нижче мінімуму.`} />
                <Task icon={<PackagePlus />} title="Потреби ремонту" text={`${requirements.length} потреб треба закрити приходом або закупкою.`} />
                <Task icon={<Truck />} title="Закупки" text={`${purchases.length} документів закупки у роботі.`} />
              </>
            ) : (
              <>
                <Task icon={<History />} title="Заказы без руху" text={`${stuckOrders.length} не змінювали статус більше 2 днів.`} />
                <Task icon={<Truck />} title="Довго чекають запчастину" text={`${orders.filter((order) => order.status === 'Очікує запчастину').length} у статусі очікування.`} />
                <Task icon={<PackageCheck />} title="Готові, але не видані" text={`${orders.filter((order) => ['Готовий до видачі', 'Не підлягає ремонту', 'Очікує клієнта', 'Очікує оплати'].includes(order.status)).length} потребують дії менеджера.`} />
                <Task icon={<Banknote />} title="Втрати" text={`Повернення ${money(returnsAmount)}, списання в ремонтах ${money(writeOffs)}.`} />
              </>
            )}
          </div>
        </div>
       
      </section>}
      {(isManager || isOwner) && <section className="content-split">
        <div className="panel">
          <div className="panel-heading"><h2>Заказы по статусах</h2><span>де зависає сервіс</span></div>
          <MiniBars rows={statusBars} />
        </div>
        <div className="panel">
          <div className="panel-heading"><h2>Склад і оборачиваемість</h2><span>{minStock.length} критичних</span></div>
          <div className="task-list">
            {topProducts.slice(0, 4).map((item) => <Task key={item.product.id} icon={<Archive />} title={item.product.name} text={`Продано ${item.sold} шт. · доступно ${item.availableQty} шт. · мінімум ${item.product.min} шт.`} />)}
          </div>
        </div>
      </section>}
      {isOwner && <section className="content-split">
        <div className="panel">
          <div className="panel-heading"><h2>Ефективність менеджерів</h2><span>{managerRows.length} співробітників</span></div>
          <div className="task-list">
            {managerRows.map((row) => <Task key={row.name} icon={<Users />} title={row.name} text={`Прийнято ${row.orders}, видано ${row.closed}, оплат ${row.paymentsCount}, середній чек ${money(row.avg)}, борги ${money(row.debt)}, акти не повернені ${row.documentsCount}.`} />)}
          </div>
        </div>
        <div className="panel">
          <div className="panel-heading"><h2>Ефективність інженерів</h2><span>{engineerRows.length} майстрів</span></div>
          <div className="task-list">
            {engineerRows.map((row) => <Task key={row.name} icon={<Wrench />} title={row.name} text={`Заказів ${row.orders}, завершено ${row.finished}, середній час ${row.avgDays} дн., без руху ${row.stuck}, повтори ${row.repeated}, заробіток ${money(row.earning)}, прибуток компанії ${money(row.companyProfit)}.`} />)}
          </div>
        </div>
      </section>}
      {isEngineer && <section className="panel">
        <div className="panel-heading"><h2>Мої найближчі дії</h2><span>{roleOrders.length} заказів</span></div>
        <OrderSummaryList orders={roleOrders.slice(0, 5)} />
      </section>} {(isManager || isAccountant || isOwner) && <section className="content-split">
        <div className="panel">
          <div className="panel-heading"><h2>Проблемні заказы</h2><span>{problemOrders.length}</span></div>
          <OrderSummaryList orders={problemOrders.length ? problemOrders.slice(0, 4) : orders.slice(0, 3)} />
        </div>
        <div className="panel">
          <div className="panel-heading"><h2>Сигнали керівнику</h2><span>автоконтроль</span></div>
          <div className="task-list">
            <Task icon={<Archive />} title="Критичний склад" text={`${minStock.length} товарів нижче мінімуму.`} />
            <Task icon={<Banknote />} title="Каса і термінал" text={`Готівка, картка і безнал контролюються окремо в розділі Оплати / Каса.`} />
            <Task icon={<History />} title="Остання дія" text={actionLogs[0]?.comment ?? 'Дій ще немає.'} />
          </div>
        </div>
      </section>}
     
     
    </div>
  );
}

function AutoControlPanel({
  orders,
  sales,
  users,
  title,
}: {
  orders: ServiceOrder[];
  sales: Sale[];
  users: User[];
  title: string;
}) {
  const todayKey = extractDayKey(today);
  const yesterdayDate = parseDateTime(today);
  const yesterday = yesterdayDate ? new Date(yesterdayDate.getTime() - 24 * 60 * 60 * 1000) : null;
  const yesterdayKey = yesterday ? `${String(yesterday.getDate()).padStart(2, '0')}.${String(yesterday.getMonth() + 1).padStart(2, '0')}.${yesterday.getFullYear()}` : '';

  const alertRows = [
    ...orders
      .filter((order) => !order.engineer)
      .map((order) => ({ id: `no-engineer-${order.id}`, orderId: order.id, tone: 'red' as const, reason: 'Без інженера' })),
    ...orders
      .filter((order) => simpleRepairStatus(order.status) === 'Готово' && orderTotals(order).debt > 0)
      .map((order) => ({ id: `ready-unpaid-${order.id}`, orderId: order.id, tone: 'yellow' as const, reason: 'Готово, але без оплати' })),
    ...orders
      .filter((order) => !['Видано', 'Закрито', 'Скасовано'].includes(order.status) && hoursSince(order.statusChangedAt) >= 48)
      .map((order) => ({ id: `stuck-${order.id}`, orderId: order.id, tone: hoursSince(order.statusChangedAt) >= 72 ? 'red' as const : 'yellow' as const, reason: `Без руху ${Math.floor(hoursSince(order.statusChangedAt) / 24)} дн.` })),
  ]
    .filter((item, index, list) => list.findIndex((row) => row.orderId === item.orderId) === index)
    .slice(0, 8);

  const engineerLoadRows = users
    .filter((user) => user.role === 'Інженер')
    .map((user) => {
      const activeCount = orders.filter((order) => order.engineer === user.name && !['Видано', 'Закрито', 'Скасовано'].includes(order.status)).length;
      const tone = activeCount >= 10 ? 'red' : activeCount >= 6 ? 'yellow' : 'green';
      return { name: user.name, count: activeCount, tone };
    })
    .sort((a, b) => b.count - a.count || a.name.localeCompare(b.name));

  const allPayments = [
    ...orders.flatMap((order) => order.payments
      .filter((payment) => paymentCountsAsApplied(payment) && payment.amount > 0)
      .map((payment) => ({ id: payment.id, date: payment.date, amount: payment.amount, method: payment.method }))),
    ...sales.flatMap((sale) => sale.payments
      .filter((payment) => paymentCountsAsApplied(payment) && payment.amount > 0)
      .map((payment) => ({ id: payment.id, date: payment.date, amount: payment.amount, method: payment.method }))),
  ];
  const todayPayments = allPayments.filter((payment) => extractDayKey(payment.date) === todayKey);
  const yesterdayPayments = allPayments.filter((payment) => extractDayKey(payment.date) === yesterdayKey);
  const todayPaymentsTotal = todayPayments.reduce((sum, payment) => sum + payment.amount, 0);
  const yesterdayPaymentsTotal = yesterdayPayments.reduce((sum, payment) => sum + payment.amount, 0);
  const todayDelta = todayPaymentsTotal - yesterdayPaymentsTotal;
  const todayDeltaLabel = todayDelta === 0 ? 'без змін до вчора' : `${todayDelta > 0 ? '+' : ''}${money(todayDelta)} до вчора`;

  const stuckOrders = orders
    .filter((order) => !['Видано', 'Закрито', 'Скасовано'].includes(order.status) && hoursSince(order.statusChangedAt) >= 48)
    .map((order) => ({ id: order.id, client: order.client, days: Math.floor(hoursSince(order.statusChangedAt) / 24) }))
    .sort((a, b) => b.days - a.days)
    .slice(0, 8);

  return (
    <section className="panel auto-control-panel">
      <div className="panel-heading">
        <h2>{title}</h2>
        <span>оновлюється автоматично</span>
      </div>

      <div className="auto-control-grid">
        <section className="auto-control-card">
          <div className="panel-heading">
            <h2>Алерти</h2>
            <span>{alertRows.length}</span>
          </div>
          <div className="auto-control-list">
            {alertRows.map((item) => (
              <div key={item.id} className={item.tone === 'red' ? 'auto-control-alert auto-control-alert-red' : 'auto-control-alert auto-control-alert-yellow'}>
                <strong>{item.orderId}</strong>
                <span>{item.reason}</span>
              </div>
            ))}
            {alertRows.length === 0 && <div className="empty-state">Критичних сигналів немає.</div>}
          </div>
        </section>

        <section className="auto-control-card">
          <div className="panel-heading">
            <h2>Завантаження інженерів</h2>
            <span>{engineerLoadRows.length}</span>
          </div>
          <div className="auto-control-list">
            {engineerLoadRows.map((row) => (
              <div key={row.name} className={`auto-control-load auto-control-load-${row.tone}`}>
                <span>{row.name}</span>
                <strong>{row.count}</strong>
              </div>
            ))}
            {engineerLoadRows.length === 0 && <div className="empty-state">Інженерів поки немає.</div>}
          </div>
        </section>

        <section className="auto-control-card">
          <div className="panel-heading">
            <h2>Фінанси сьогодні</h2>
            <span>{todayPayments.length} оплат</span>
          </div>
          <div className="auto-control-finance">
            <strong>{money(todayPaymentsTotal)}</strong>
            <span>{todayDeltaLabel}</span>
          </div>
        </section>

        <section className="auto-control-card">
          <div className="panel-heading">
            <h2>Зависші замовлення</h2>
            <span>{stuckOrders.length}</span>
          </div>
          <div className="auto-control-list">
            {stuckOrders.map((item) => (
              <div key={item.id} className="auto-control-stuck">
                <strong>{item.id}</strong>
                <span>{item.client}</span>
                <em>{item.days} дн.</em>
              </div>
            ))}
            {stuckOrders.length === 0 && <div className="empty-state">Без завислих замовлень.</div>}
          </div>
        </section>
      </div>
    </section>
  );
}

function EmployeeControlPage({ orders, users }: { orders: ServiceOrder[]; users: User[] }) {
  const [sortBy, setSortBy] = useState<'load' | 'profit' | 'speed'>('load');
  const [selectedEngineer, setSelectedEngineer] = useState('');
  const todayDate = parseDateTime(today);
  const currentMonthKey = todayDate ? `${todayDate.getFullYear()}-${todayDate.getMonth()}` : '';
  const isInCurrentMonth = (dateText: string) => {
    const parsed = parseDateTime(dateText);
    if (!parsed || !currentMonthKey) return false;
    return `${parsed.getFullYear()}-${parsed.getMonth()}` === currentMonthKey;
  };

  const engineerRows = users
    .filter((user) => user.role === 'Інженер')
    .map((user) => {
      const engineerOrders = orders.filter((order) => order.engineer === user.name);
      const inWorkOrders = engineerOrders.filter((order) => simpleRepairStatus(order.status) === 'В ремонті');
      const readyOrders = engineerOrders.filter((order) => simpleRepairStatus(order.status) === 'Готово');
      const issuedOrders = engineerOrders.filter((order) => simpleRepairStatus(order.status) === 'Видано');
      const avgRepairDays = engineerOrders.length
        ? engineerOrders.reduce((sum, order) => sum + daysSince(order.intakeDate), 0) / engineerOrders.length
        : 0;
      const profit = engineerOrders.reduce((sum, order) => sum + orderTotals(order).finalProfit, 0);
      const monthlyDoneOrders = issuedOrders.filter((order) => isInCurrentMonth(order.statusChangedAt));
      const monthlyDoneCount = monthlyDoneOrders.length;
      const monthlyEarned = engineerOrders.reduce((sum, order) => (
        sum + order.works.reduce((inner, work) => {
          const workEngineer = work.engineer ?? order.engineer;
          if (workEngineer !== user.name) return inner;
          return inner + payrollForWork(order, work, payrollRules.find((rule) => rule.employee === workEngineer));
        }, 0)
      ), 0);
      const averageCheck = engineerOrders.length
        ? Math.round(engineerOrders.reduce((sum, order) => sum + orderTotals(order).paid, 0) / engineerOrders.length)
        : 0;
      const hasStuck = inWorkOrders.some((order) => daysSince(order.statusChangedAt) > 3);
      const loadState = hasStuck ? 'critical' : inWorkOrders.length >= 8 ? 'warning' : 'ok';
      return {
        name: user.name,
        inWork: inWorkOrders.length,
        ready: readyOrders.length,
        issued: issuedOrders.length,
        avgRepairDays,
        profit,
        monthlyDoneCount,
        monthlyEarned,
        averageCheck,
        loadState,
        orders: engineerOrders
          .map((order) => ({
            id: order.id,
            client: order.client,
            status: managerOrderStatusLabel(order.status),
            daysInWork: daysSince(order.statusChangedAt),
            isStuck: daysSince(order.statusChangedAt) > 3,
          }))
          .sort((a, b) => b.daysInWork - a.daysInWork),
      };
    });

  const sortedRows = [...engineerRows].sort((a, b) => {
    if (sortBy === 'profit') return b.profit - a.profit || b.inWork - a.inWork;
    if (sortBy === 'speed') return a.avgRepairDays - b.avgRepairDays || b.issued - a.issued;
    return b.inWork - a.inWork || b.profit - a.profit;
  });

  const activeEngineer = sortedRows.find((row) => row.name === selectedEngineer) ?? sortedRows[0];

  return (
    <div className="page-grid employee-control-page">
      <PageTitle eyebrow="Контроль" title="Контроль співробітників" text="Ефективність інженерів, завантаження і замовлення без руху." />

      <section className="panel employee-control-panel">
        <div className="panel-heading">
          <h2>Інженери</h2>
          <div className="employee-control-sort">
            <button type="button" className={sortBy === 'load' ? 'primary' : ''} onClick={() => setSortBy('load')}>Завантаження</button>
            <button type="button" className={sortBy === 'profit' ? 'primary' : ''} onClick={() => setSortBy('profit')}>Прибуток</button>
            <button type="button" className={sortBy === 'speed' ? 'primary' : ''} onClick={() => setSortBy('speed')}>Швидкість</button>
          </div>
        </div>

        <div className="employee-control-table">
          <div className="employee-control-head">
            <span>Інженер</span>
            <span>В роботі</span>
            <span>Готово</span>
            <span>Видано</span>
            <span>Середній час</span>
            <span>Прибуток</span>
          </div>
          {sortedRows.map((row) => (
            <button
              key={row.name}
              type="button"
              className={`employee-control-row employee-control-${row.loadState}${activeEngineer?.name === row.name ? ' is-active' : ''}`}
              onClick={() => setSelectedEngineer(row.name)}
            >
              <span>{row.name}</span>
              <span>{row.inWork}</span>
              <span>{row.ready}</span>
              <span>{row.issued}</span>
              <span>{row.avgRepairDays.toFixed(1)} дн.</span>
              <strong>{money(row.profit)}</strong>
            </button>
          ))}
          {sortedRows.length === 0 && <div className="empty-state">Інженерів поки немає.</div>}
        </div>
      </section>

      {activeEngineer && (
        <section className="employee-control-grid">
          <section className="panel employee-control-panel">
            <div className="panel-heading">
              <h2>{activeEngineer.name}</h2>
              <span>{activeEngineer.loadState === 'critical' ? '🔴 ризик' : activeEngineer.loadState === 'warning' ? '🟠 перевантаження' : '🟢 норма'}</span>
            </div>
            <div className="employee-control-kpis">
              <article className="employee-control-kpi">
                <span>За місяць</span>
                <strong>{activeEngineer.monthlyDoneCount}</strong>
              </article>
              <article className="employee-control-kpi">
                <span>Заробив</span>
                <strong>{money(activeEngineer.monthlyEarned)}</strong>
              </article>
              <article className="employee-control-kpi">
                <span>Середній чек</span>
                <strong>{money(activeEngineer.averageCheck)}</strong>
              </article>
              <article className="employee-control-kpi">
                <span>Середній ремонт</span>
                <strong>{activeEngineer.avgRepairDays.toFixed(1)} дн.</strong>
              </article>
            </div>
          </section>

          <section className="panel employee-control-panel">
            <div className="panel-heading">
              <h2>Замовлення інженера</h2>
              <span>{activeEngineer.orders.length}</span>
            </div>
            <div className="employee-control-orders">
              <div className="employee-control-orders-head">
                <span>Номер</span>
                <span>Клієнт</span>
                <span>Статус</span>
                <span>Днів</span>
              </div>
              {activeEngineer.orders.map((order) => (
                <div key={order.id} className={order.isStuck ? 'employee-control-order-row is-stuck' : 'employee-control-order-row'}>
                  <span>{order.id}</span>
                  <span>{order.client}</span>
                  <span>{order.status}</span>
                  <strong>{order.daysInWork}</strong>
                </div>
              ))}
              {activeEngineer.orders.length === 0 && <div className="empty-state">У цього інженера ще немає замовлень.</div>}
            </div>
          </section>
        </section>
      )}
    </div>
  );
}

function ContractsPage({
  contracts,
  contractActs,
  orders,
  documents,
  createContract,
  createContractAct,
  createInvoiceForContractAct,
  closeContract,
  readOnly,
}: {
  contracts: ContractRecord[];
  contractActs: ContractActRecord[];
  orders: ServiceOrder[];
  documents: PrintDocument[];
  createContract: (payload: { client: string; amount: string; startDate: string; endDate: string }) => boolean;
  createContractAct: (contractId: string, orderIds: string[]) => boolean;
  createInvoiceForContractAct: (actId: string) => void;
  closeContract: (contractId: string) => void;
  readOnly: boolean;
}) {
  const [client, setClient] = useState('');
  const [amount, setAmount] = useState('');
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [draftActContractId, setDraftActContractId] = useState('');
  const [draftActOrderIds, setDraftActOrderIds] = useState<string[]>([]);

  const sortedContracts = [...contracts].sort((a, b) => (parseDateTime(b.createdAt)?.getTime() ?? 0) - (parseDateTime(a.createdAt)?.getTime() ?? 0));

  return (
    <div className="page-grid contracts-page">
      <PageTitle eyebrow="Договори" title="Робота по договорах" text="Окремий контур для бюджетних клієнтів: ліміт, використання, документи і закриття." />

      {!readOnly && (
      <section className="panel contracts-create">
        <div className="panel-heading">
          <h2>Новий договір</h2>
          <span>{sortedContracts.length} всього</span>
        </div>
        <div className="table">
          <label>
            Клієнт
            <input value={client} onChange={(event) => setClient(event.target.value)} placeholder="ТОВ / установа" />
          </label>
          <label>
            Сума договору
            <input value={amount} onChange={(event) => setAmount(event.target.value)} type="number" min={0} placeholder="0" />
          </label>
          <label>
            Дата початку
            <input value={startDate} onChange={(event) => setStartDate(event.target.value)} type="date" />
          </label>
          <label>
            Дата завершення
            <input value={endDate} onChange={(event) => setEndDate(event.target.value)} type="date" />
          </label>
        </div>
        <div className="action-row">
          <button
            type="button"
            className="submit-button"
            onClick={() => {
              const created = createContract({ client, amount, startDate, endDate });
              if (created) {
                setClient('');
                setAmount('');
                setStartDate('');
                setEndDate('');
              }
            }}
          >
            Створити договір
          </button>
        </div>
      </section>
      )}

      <section className="contracts-grid">
        {sortedContracts.map((contract) => {
          const usage = contractUsage(contract, orders);
          const closedAmount = contractClosedAmount(contract, contractActs);
          const contractOrders = usage.contractOrders.sort((a, b) => (parseDateTime(b.intakeDate)?.getTime() ?? 0) - (parseDateTime(a.intakeDate)?.getTime() ?? 0));
          const draftableOrders = contractOrders.filter((order) => order.contractAccountedAt && !order.contractActId);
          const contractRelatedActs = contractActs.filter((act) => act.contractId === contract.id).sort((a, b) => (parseDateTime(b.date)?.getTime() ?? 0) - (parseDateTime(a.date)?.getTime() ?? 0));
          const actDocuments = documents.filter((document) => document.entityType === 'act' && contractRelatedActs.some((act) => act.id === document.entityId));
          const isDraftOpen = draftActContractId === contract.id;
          const selectedDraftOrders = draftableOrders.filter((order) => draftActOrderIds.includes(order.id));
          const draftActTotal = selectedDraftOrders.reduce((sum, order) => sum + contractOrderAmount(order), 0);
          return (
            <article key={contract.id} className="panel contracts-card">
              <div className="panel-heading">
                <div>
                  <h2>{contract.id}</h2>
                  <span>{contract.client}</span>
                </div>
                <span className={contract.status === 'Активний' ? 'tag tag-green' : 'tag tag-gray'}>{contract.status}</span>
              </div>

              <div className="contracts-kpis">
                <div>
                  <span>Ліміт</span>
                  <strong>{money(contract.amount)}</strong>
                </div>
                <div>
                  <span>Використано</span>
                  <strong>{money(usage.used)}</strong>
                </div>
                <div>
                  <span>Закрито актами</span>
                  <strong>{money(closedAmount)}</strong>
                </div>
                <div>
                  <span>Залишок</span>
                  <strong>{money(usage.remaining)}</strong>
                </div>
              </div>

              <div className="client-meta">
                <span>{contract.startDate} - {contract.endDate}</span>
                <span>{contract.closedAt ? `Закрито ${contract.closedAt}` : `Створено ${contract.createdAt}`}</span>
              </div>

              {!readOnly && (
                <div className="action-row">
                  <button
                    type="button"
                    onClick={() => {
                      setDraftActContractId((current) => current === contract.id ? '' : contract.id);
                      setDraftActOrderIds([]);
                    }}
                  >
                    Сформувати акт
                  </button>
                  <button type="button" className="submit-button" onClick={() => closeContract(contract.id)} disabled={contract.status === 'Закритий'}>Закрити договір</button>
                </div>
              )}

              {!readOnly && isDraftOpen && (
                <div className="contracts-act-builder">
                  <div className="panel-heading">
                    <h2>Новий акт</h2>
                    <span>{draftableOrders.length} доступно</span>
                  </div>
                  <div className="contracts-act-builder-list">
                    {draftableOrders.map((order) => (
                      <label key={order.id} className="contracts-act-order">
                        <input
                          type="checkbox"
                          checked={draftActOrderIds.includes(order.id)}
                          onChange={(event) => {
                            setDraftActOrderIds((current) => (
                              event.target.checked
                                ? [...current, order.id]
                                : current.filter((id) => id !== order.id)
                            ));
                          }}
                        />
                        <span>{order.id}</span>
                        <span>{order.client}</span>
                        <strong>{money(contractOrderAmount(order))}</strong>
                      </label>
                    ))}
                    {draftableOrders.length === 0 && <div className="empty-state">Немає замовлень, доступних для нового акта.</div>}
                  </div>
                  <div className="client-meta">
                    <span>Сума акта</span>
                    <strong>{money(draftActTotal)}</strong>
                  </div>
                  <div className="action-row">
                    <button
                      type="button"
                      className="submit-button"
                      onClick={() => {
                        const created = createContractAct(contract.id, draftActOrderIds);
                        if (created) {
                          setDraftActOrderIds([]);
                          setDraftActContractId('');
                        }
                      }}
                      disabled={draftActOrderIds.length === 0}
                    >
                      Створити акт
                    </button>
                    <button type="button" onClick={() => { setDraftActOrderIds([]); setDraftActContractId(''); }}>Скасувати</button>
                  </div>
                </div>
              )}

              <div className="contracts-docs">
                <strong>Акти</strong>
                <div className="task-list">
                  {contractRelatedActs.length > 0
                    ? contractRelatedActs.map((act) => {
                      const actOrders = orders.filter((order) => act.orderIds.includes(order.id));
                      const actInvoice = actDocuments.find((document) => document.entityId === act.id && document.kind === 'Рахунок на оплату');
                      return (
                        <div key={act.id} className="contracts-act-item">
                          <div className="contracts-act-head">
                            <strong>{act.id}</strong>
                            <span className={actPaymentStatusClass(act.status)}>{act.status}</span>
                          </div>
                          <div className="client-meta">
                            <span>{act.date}</span>
                            <strong>{money(act.amount)}</strong>
                          </div>
                          <div className="client-meta">
                            <span>Сума акта: {money(act.amount)}</span>
                            <span>Оплачено: {money(act.paidAmount)}</span>
                          </div>
                          <div className="client-meta">
                            <span>Залишок: {money(act.remainingAmount)}</span>
                            <span>{actInvoice ? 'Оплата тільки в межах залишку акта' : 'Рахунок ще не сформовано'}</span>
                          </div>
                          <div className="contracts-act-orders">
                            {actOrders.map((order) => <span key={order.id}>{order.id}</span>)}
                          </div>
                          {!readOnly && (
                            <div className="action-row">
                              <button type="button" onClick={() => createInvoiceForContractAct(act.id)} disabled={Boolean(actInvoice)}>Сформувати рахунок</button>
                            </div>
                          )}
                          {actInvoice && (
                            <div className="client-meta">
                              <span>{actInvoice.number}</span>
                              <span>{actInvoice.status}</span>
                            </div>
                          )}
                        </div>
                      );
                    })
                    : <div className="empty-state">Актів поки немає.</div>}
                </div>
              </div>

              <div className="contracts-orders">
                <div className="panel-heading">
                  <h2>Замовлення</h2>
                  <span>{contractOrders.length}</span>
                </div>
                <div className="table">
                  <div className="table-row table-head">
                    <span>Замовлення</span>
                    <span>Клієнт</span>
                    <span>Статус</span>
                    <span>Сума</span>
                    <span>Акт</span>
                  </div>
                  {contractOrders.map((order) => (
                    <div key={order.id} className="table-row">
                      <span>{order.id}</span>
                      <span>{order.client}</span>
                      <span>{managerOrderStatusLabel(order.status)}</span>
                      <span>{money(contractOrderAmount(order))}</span>
                      <span>{order.contractActId ?? 'Не закрито'}</span>
                    </div>
                  ))}
                </div>
                {contractOrders.length === 0 && <div className="empty-state">Ще немає замовлень по цьому договору.</div>}
              </div>
            </article>
          );
        })}
        {sortedContracts.length === 0 && <div className="panel empty-state">Договорів поки немає.</div>}
      </section>
    </div>
  );
}

function OrdersPage(props: {
  orders: ServiceOrder[];
  allOrders: ServiceOrder[];
  allRoleOrders?: ServiceOrder[];
  selectedOrder: ServiceOrder;
  products: Product[];
  selectedOrderId: string;
  setSelectedOrderId: (id: string) => void;
  selectedProductId: string;
  setSelectedProductId: (id: string) => void;
  qty: number;
  setQty: (qty: number) => void;
  quickPhone: string;
  quickClientDebtWarning: string;
  quickClientName: string;
  quickDevice: string;
  quickSerial: string;
  quickProblem: string;
  quickAppearance: string;
  quickEstimatedAmount: string;
  quickEngineerId: string;
  quickContractId: string;
  quickLocationCode: string;
  quickComment: string;
  users: User[];
  contracts: ContractRecord[];
  contractActs: ContractActRecord[];
  setQuickPhone: (phone: string) => void;
  setQuickClientName: (name: string) => void;
  setQuickDevice: (device: string) => void;
  setQuickSerial: (serial: string) => void;
  setQuickProblem: (problem: string) => void;
  setQuickAppearance: (appearance: string) => void;
  setQuickEstimatedAmount: (amount: string) => void;
  setQuickEngineerId: (id: string) => void;
  setQuickContractId: (id: string) => void;
  setQuickLocationCode: (code: string) => void;
  setQuickComment: (comment: string) => void;
  createQuickOrder: () => void;
  patchSimpleManagerOrder: (orderId: string, patch: Partial<ServiceOrder>) => void;
  updateSimpleManagerOrderStatus: (orderId: string, nextStatus: 'В ремонті' | 'Готово' | 'Видано') => void;
  editSimpleManagerOrder: (orderId: string, payload: { phone: string; client: string; device: string; issue: string; estimatedAmount: string; engineerId: string }) => void;
  addSimpleManagerOrderPart: (orderId: string, productId: string, qtyToAdd: number) => void;
  removeSimpleManagerOrderPart: (orderId: string, partId: string, reason?: string, comment?: string) => void;
  acceptSimpleManagerPayment: (orderId: string, repairPriceInput: string, paymentType: 'наличные' | 'карта' | 'перевод', paymentKind: SimpleLedgerPaymentKind, paymentReason?: string) => string | null;
  bankImportItems?: BankImportItem[];
  confirmSimpleManagerPayment: (paymentId: string) => void;
  accountSimpleManagerOrderToContract: (orderId: string) => void;
  warehouseLocations: WarehouseLocation[];
  customerList: ClientRecord[];
  addPartToRepair: () => void;
  orderPart: (order: ServiceOrder, part: OrderPart) => void;
  reserveArrived: (order: ServiceOrder, part: OrderPart) => void;
  issueToEngineer: (order: ServiceOrder, part: OrderPart) => void;
  markInstalled: (order: ServiceOrder, part: OrderPart) => void;
  returnServicePart: (order: ServiceOrder, part: OrderPart, destination: 'На склад' | 'Брак') => void;
  addOrderPayment: (order: ServiceOrder, amount: number, method: PaymentMethod, comment: string) => void;
  changeOrderStatus: (order: ServiceOrder, nextStatus: OrderStatus, comment: string) => void;
  acceptOrderWork: (order: ServiceOrder) => void;
  returnOrderToCellReady: (order: ServiceOrder) => void;
  reassignEngineer: (order: ServiceOrder) => void;
  closeOrder: (order: ServiceOrder) => void;
  issueReadyOrder: (order: ServiceOrder) => void;
  oneClickManagerIssue: (order: ServiceOrder, payment?: { method: PaymentMethod; amount?: number; partial?: boolean }) => void;
  transferOrderToBas: (order: ServiceOrder) => void;
  printDocument: (kind: DocumentKind, entityType: PrintDocument['entityType'], entityId: string, clientOrSupplier: string) => void;
  createServiceOrderDocument: (kind: ServiceOrderDocumentKind, order: ServiceOrder) => PrintDocument | undefined;
  printServiceOrderDocument: (kind: ServiceOrderDocumentKind, order: ServiceOrder) => void;
  documents: PrintDocument[];
  taxInvoices: TaxInvoice[];
  notifications: ClientNotification[];
  allNotifications: ClientNotification[];
  orderUnits: OrderUnit[];
  orderMovementLogs: OrderMovementLog[];
  moveOrder: (orderId: string, newLocation: string) => void;
  sendClientNotification: (order: ServiceOrder, event: NotificationEvent, preferredChannel?: NotificationChannel, manual?: boolean, messageOverride?: string) => void;
  createNotificationDraft: (order: ServiceOrder, event: NotificationEvent, channel: NotificationChannel, text: string, status: NotificationStatus) => ClientNotification;
  approval?: RepairApproval;
  sendRepairApproval: (order: ServiceOrder, channel?: NotificationChannel, extra?: { description?: string; amount?: number; comment?: string }) => void;
  recordApprovalResponse: (approvalId: string, accepted: boolean, manual?: boolean) => void;
  markApprovalNoAnswer: (approvalId: string) => void;
  logRiskConfirmation: (order: ServiceOrder, action: string, risks: string[]) => void;
  ensureOrderDocumentRecord: (kind: 'Рахунок на оплату' | 'Акт надання послуг' | 'Наряд на ремонт' | 'Гарантійний талон', order: ServiceOrder) => void;
  logOrderDocumentPrint: (order: ServiceOrder, kind: 'Рахунок на оплату' | 'Акт надання послуг' | 'Наряд на ремонт' | 'Гарантійний талон') => void;
  signOrderAct: (order: ServiceOrder) => void;
  createTaxInvoiceForOrder: (orderId: string) => void;
  registerTaxInvoice: (orderId: string) => void;
  markEngineerWorkCompleted: (order: ServiceOrder) => void;
  cancelOrder: (order: ServiceOrder, reason?: string, comment?: string) => void;
  refundOrder: (order: ServiceOrder, amount?: number, reason?: string, comment?: string) => void;
  cancelOrderAct: (order: ServiceOrder) => void;
  reopenOrder: (order: ServiceOrder, reason?: string, comment?: string) => void;
  openClientRecord: (phone: string, search?: string) => void;
  orderVersions: OrderVersion[];
  suggestedDocumentAction: SuggestedDocumentAction | null;
  clearSuggestedDocumentAction: () => void;
  dashboardFocus: { orderId: string; target: DashboardFocusTarget } | null;
  clearDashboardFocus: () => void;
  notifyUser: (message: string) => void;
  canDo: (permission: Permission) => boolean;
  showCost: boolean;
  activeUser: User;
}) {
  const isEngineer = props.activeUser.role === 'Інженер';
  const isManager = props.activeUser.role === 'Менеджер';
  const isSupervisor = props.activeUser.role === 'Руководитель';
  const deviceIcon = (device: string) => {
    const value = device.toLowerCase();
    if (value.includes('printer') || value.includes('laserjet') || value.includes('принтер') || value.includes('мфу')) return '🖨';
    if (value.includes('ноут') || value.includes('laptop') || value.includes('thinkpad') || value.includes('macbook')) return '💻';
    if (value.includes('телефон') || value.includes('iphone') || value.includes('samsung') || value.includes('смартф')) return '📱';
    return '🧾';
  };
  const [managerSearch, setManagerSearch] = useState('');
  const [managerFilter, setManagerFilter] = useState<'all' | 'Прийнято' | 'В ремонті' | 'Готово' | 'Видано' | 'Борг' | 'Очікує оплату'>('all');
  const [showManagerCreateForm, setShowManagerCreateForm] = useState(false);
  const [addingPartOrderId, setAddingPartOrderId] = useState('');
  const [partSearch, setPartSearch] = useState('');
  const [selectedPartProductId, setSelectedPartProductId] = useState('');
  const [selectedPartQty, setSelectedPartQty] = useState('1');
  const [managerActiveOrderId, setManagerActiveOrderId] = useState('');
  const [isManagerOrderDetailOpen, setIsManagerOrderDetailOpen] = useState(false);
  const [managerPulseOrderId, setManagerPulseOrderId] = useState('');
  const [managerPaymentDrafts, setManagerPaymentDrafts] = useState<Record<string, { amount: string; type: SimpleLedgerPaymentKind; method: 'наличные' | 'карта' | 'перевод'; reason: string }>>({});
  const [managerPaymentModal, setManagerPaymentModal] = useState<null | { orderId: string; amount: string; due: number; method: 'наличные' | 'карта' | 'перевод' }>(null);
  const [managerIssueModal, setManagerIssueModal] = useState<null | { orderId: string }>(null);
  const [managerPostPaymentOrderId, setManagerPostPaymentOrderId] = useState('');
  const [managerExceptionDraft, setManagerExceptionDraft] = useState<{ orderId: string; mode: 'cancel' | 'reopen' | 'refund' | 'part-return'; reason: string; comment: string; amount: string; partId?: string } | null>(null);
  const managerPaymentAmountInputRef = useRef<HTMLInputElement | null>(null);
  const roleOrders = props.allRoleOrders ?? props.orders;
  const managerSourceOrders = props.allRoleOrders ?? props.allOrders;
  const managerEngineers = props.users.filter((user) => user.role === 'Інженер');
  const managerContracts = props.contracts.filter((contract) => contract.status === 'Активний');
  const managerAvailableProducts = props.products.filter((product) => available(product) > 0);
  const suggestedManagerEngineer = managerEngineers[0];
  const managerOrderPriorityScore = (order: ServiceOrder) => {
    const total = Math.max(order.repairPrice ?? order.estimatedAmount ?? orderTotals(order).total, 0);
    const paid = order.payments.filter(paymentCountsAsApplied).reduce((sum, payment) => sum + payment.amount, 0);
    const remaining = Math.max(total - paid, 0);
    let score = 0;
    if (remaining > 0 && order.status === 'Видано') score += 500;
    if (remaining > 0 && simpleRepairStatus(order.status) === 'Готово' && order.status !== 'Видано') score += 400;
    if (!order.engineer) score += 300;
    if (remaining > 0) score += 120;
    if (simpleRepairStatus(order.status) === 'В ремонті') score += 80;
    if (simpleRepairStatus(order.status) === 'Прийнято') score += 40;
    score += Math.min(daysSince(order.statusChangedAt), 60);
    return score;
  };
  const managerFilterMatch = (order: ServiceOrder) => {
    const total = Math.max(order.repairPrice ?? order.estimatedAmount ?? orderTotals(order).total, 0);
    const paid = order.payments.filter(paymentCountsAsApplied).reduce((sum, payment) => sum + payment.amount, 0);
    const remaining = Math.max(total - paid, 0);
    if (managerFilter === 'all') return true;
    if (managerFilter === 'Борг') return remaining > 0 && (Boolean(order.legalEntity) || Boolean(order.contractId) || simpleRepairStatus(order.status) === 'Видано');
    if (managerFilter === 'Очікує оплату') return remaining > 0 && simpleRepairStatus(order.status) === 'Готово' && order.status !== 'Видано';
    return simpleRepairStatus(order.status) === managerFilter;
  };
  const managerVisibleOrders = [...managerSourceOrders]
    .filter((order) => !['Скасовано', 'Закрито'].includes(order.status))
    .filter((order) => {
      const needle = managerSearch.trim();
      if (!needle) return true;
      return matchesOrderSearch(order, needle);
    })
    .filter((order) => managerFilterMatch(order))
    .sort((a, b) => (parseDateTime(b.createdAt ?? b.intakeDate)?.getTime() ?? 0)
      - (parseDateTime(a.createdAt ?? a.intakeDate)?.getTime() ?? 0));
  const activeManagerOrder = managerVisibleOrders.find((order) => order.id === managerActiveOrderId) ?? managerVisibleOrders[0];
  const managerAlerts = {
    unpaid: managerSourceOrders.filter((order) => {
      if (['Скасовано', 'Закрито'].includes(order.status)) return false;
      const total = Math.max(order.repairPrice ?? order.estimatedAmount ?? orderTotals(order).total, 0);
      const paid = order.payments.filter(paymentCountsAsApplied).reduce((sum, payment) => sum + payment.amount, 0);
      return Math.max(total - paid, 0) > 0 && simpleRepairStatus(order.status) === 'Готово' && order.status !== 'Видано';
    }).length,
    debt: managerSourceOrders.filter((order) => {
      if (['Скасовано', 'Закрито'].includes(order.status)) return false;
      const total = Math.max(order.repairPrice ?? order.estimatedAmount ?? orderTotals(order).total, 0);
      const paid = order.payments.filter(paymentCountsAsApplied).reduce((sum, payment) => sum + payment.amount, 0);
      return Math.max(total - paid, 0) > 0 && (Boolean(order.legalEntity) || Boolean(order.contractId) || simpleRepairStatus(order.status) === 'Видано');
    }).length,
    noEngineer: managerSourceOrders.filter((order) => !order.engineer && !['Скасовано', 'Закрито'].includes(order.status)).length,
    readyWaiting: managerSourceOrders.filter((order) => !['Скасовано', 'Закрито'].includes(order.status) && simpleRepairStatus(order.status) === 'Готово' && order.status !== 'Видано').length,
  };

  const pulseManagerOrder = (orderId: string) => {
    setManagerPulseOrderId(orderId);
    window.setTimeout(() => setManagerPulseOrderId((current) => current === orderId ? '' : current), 700);
  };

  const managerPaymentDraft = (order: ServiceOrder) => managerPaymentDrafts[order.id] ?? {
    amount: String(Math.max((order.repairPrice ?? order.estimatedAmount ?? orderTotals(order).total) - order.payments.filter(paymentCountsAsApplied).reduce((sum, payment) => sum + payment.amount, 0), 0) || ''),
    type: 'оплата' as SimpleLedgerPaymentKind,
    method: order.repairPaymentMethod ?? 'наличные',
    reason: '',
  };

  useEffect(() => {
    if (!isManager) return;
    if (!managerVisibleOrders.some((order) => order.id === managerActiveOrderId)) {
      setManagerActiveOrderId(managerVisibleOrders[0]?.id ?? '');
    }
  }, [isManager, managerVisibleOrders, managerActiveOrderId]);

  useEffect(() => {
    if (!isManager) return;
    const handleManagerHotkeys = (event: KeyboardEvent) => {
      if (!activeManagerOrder) return;
      if (event.ctrlKey && event.key === 'Enter') {
        event.preventDefault();
        if (activeManagerOrder.contractId && !activeManagerOrder.contractAccountedAt) {
          props.accountSimpleManagerOrderToContract(activeManagerOrder.id);
          pulseManagerOrder(activeManagerOrder.id);
          return;
        }
        if (activeManagerOrder.repairPrice && activeManagerOrder.repairPrice > 0 && activeManagerOrder.status !== 'Оплачено' && activeManagerOrder.status !== 'Видано') {
          const draft = managerPaymentDraft(activeManagerOrder);
          props.acceptSimpleManagerPayment(activeManagerOrder.id, draft.amount || String(activeManagerOrder.repairPrice ?? ''), draft.method, draft.type);
          pulseManagerOrder(activeManagerOrder.id);
        }
        return;
      }
      if (event.key === 'Escape') {
        if (addingPartOrderId) setAddingPartOrderId('');
        if (showManagerCreateForm) setShowManagerCreateForm(false);
        return;
      }
      if (event.key === 'ArrowDown' || event.key === 'ArrowRight' || event.key === 'ArrowUp' || event.key === 'ArrowLeft') {
        const currentIndex = managerVisibleOrders.findIndex((order) => order.id === activeManagerOrder.id);
        if (currentIndex === -1) return;
        const nextIndex = event.key === 'ArrowDown' || event.key === 'ArrowRight'
          ? Math.min(managerVisibleOrders.length - 1, currentIndex + 1)
          : Math.max(0, currentIndex - 1);
        if (nextIndex !== currentIndex) {
          event.preventDefault();
          setManagerActiveOrderId(managerVisibleOrders[nextIndex].id);
        }
      }
    };
    window.addEventListener('keydown', handleManagerHotkeys);
    return () => window.removeEventListener('keydown', handleManagerHotkeys);
  }, [activeManagerOrder, addingPartOrderId, isManager, managerVisibleOrders, props, showManagerCreateForm]);

  useEffect(() => {
    if (!isManager) return;
    if (!managerPaymentModal && !managerIssueModal) return;
    const handleModalEscape = (event: KeyboardEvent) => {
      if (event.key !== 'Escape') return;
      if (managerPaymentModal) {
        setManagerPaymentModal(null);
        return;
      }
      if (managerIssueModal) {
        setManagerIssueModal(null);
      }
    };
    window.addEventListener('keydown', handleModalEscape);
    return () => window.removeEventListener('keydown', handleModalEscape);
  }, [isManager, managerIssueModal, managerPaymentModal]);

  useEffect(() => {
    if (!managerPaymentModal) return;
    window.setTimeout(() => managerPaymentAmountInputRef.current?.focus(), 0);
  }, [managerPaymentModal]);

  if (isManager) {
    // Manager order card contract:
    // - payment block must stay in decision area, never below parts/history
    // - do not duplicate sum/paid/remaining in multiple zones
    // - do not mix or reorder ZONE 1 / ZONE 2 / ZONE 3 without explicit product approval
    // - any new field must live inside its own zone and must not reduce clarity of can_release
    // Review rule for any change here:
    // "Can a manager understand in 2 seconds whether the order can be released?"
    // If not, the change should be rejected.
    const runManagerStatusAction = (orderId: string, nextStatus: 'В ремонті' | 'Готово' | 'Видано') => {
      props.updateSimpleManagerOrderStatus(orderId, nextStatus);
    };

    const submitManagerOrder = () => {
      if (!props.quickClientName.trim() || !props.quickPhone.trim() || !props.quickDevice.trim() || !props.quickProblem.trim()) {
        props.notifyUser('Заповніть імʼя, телефон, пристрій і проблему.');
        return;
      }
      props.createQuickOrder();
      setManagerFilter('all');
      setManagerSearch('');
      setShowManagerCreateForm(false);
    };

    const resetManagerCreateDraft = () => {
      props.setQuickClientName('');
      props.setQuickPhone('');
      props.setQuickDevice('');
      props.setQuickSerial('');
      props.setQuickProblem('');
      props.setQuickAppearance('');
      props.setQuickEstimatedAmount('');
      props.setQuickEngineerId('');
      props.setQuickContractId('');
      props.setQuickLocationCode('');
      props.setQuickComment('');
      setManagerSearch('');
    };

    const openManagerCreateMode = () => {
      resetManagerCreateDraft();
      setIsManagerOrderDetailOpen(false);
      setShowManagerCreateForm(true);
    };

    const cancelManagerCreateMode = () => {
      resetManagerCreateDraft();
      setShowManagerCreateForm(false);
    };

    const openManagerOrderDetails = (orderId: string) => {
      setManagerActiveOrderId(orderId);
      props.setSelectedOrderId(orderId);
      setIsManagerOrderDetailOpen(true);
    };

    const submitManagerPart = (orderId: string) => {
      const qtyValue = Number(selectedPartQty);
      if (!selectedPartProductId || !Number.isFinite(qtyValue) || qtyValue <= 0) {
        props.notifyUser('Оберіть деталь і вкажіть кількість.');
        return;
      }
      props.addSimpleManagerOrderPart(orderId, selectedPartProductId, qtyValue);
      pulseManagerOrder(orderId);
      setAddingPartOrderId('');
      setPartSearch('');
      setSelectedPartProductId('');
      setSelectedPartQty('1');
    };

    const resetManagerException = () => setManagerExceptionDraft(null);

    const submitManagerException = (order: ServiceOrder) => {
      if (!managerExceptionDraft || managerExceptionDraft.orderId !== order.id) return;
      const reason = managerExceptionDraft.reason.trim();
      const comment = managerExceptionDraft.comment.trim();
      if (!reason) {
        props.notifyUser('Вкажіть причину дії.');
        return;
      }
      if (managerExceptionDraft.mode === 'cancel') {
        props.cancelOrder(order, reason, comment);
        resetManagerException();
        return;
      }
      if (managerExceptionDraft.mode === 'reopen') {
        props.reopenOrder(order, reason, comment);
        resetManagerException();
        return;
      }
      if (managerExceptionDraft.mode === 'refund') {
        const refundAmount = Number(managerExceptionDraft.amount);
        if (!Number.isFinite(refundAmount) || refundAmount <= 0) {
          props.notifyUser('Вкажіть суму повернення.');
          return;
        }
        props.refundOrder(order, refundAmount, reason, comment);
        resetManagerException();
        return;
      }
      if (managerExceptionDraft.mode === 'part-return' && managerExceptionDraft.partId) {
        props.removeSimpleManagerOrderPart(order.id, managerExceptionDraft.partId, reason, comment);
        resetManagerException();
      }
    };

    const handleManagerFieldKeyDown = (event: React.KeyboardEvent<HTMLElement>) => {
      if (event.key === 'Enter' && !event.ctrlKey && !(event.target instanceof HTMLTextAreaElement && event.shiftKey)) {
        event.preventDefault();
        (event.currentTarget as HTMLElement).blur();
      }
      if (event.key === 'Escape') {
        event.preventDefault();
        (event.currentTarget as HTMLElement).blur();
      }
    };

    const suggestManagerEngineer = () => {
      const ranked = managerEngineers
        .map((engineer) => ({
          engineer,
          activeCount: managerSourceOrders.filter((order) => order.engineer === engineer.name && !['Видано', 'Закрито', 'Скасовано'].includes(order.status)).length,
        }))
        .sort((a, b) => a.activeCount - b.activeCount || a.engineer.name.localeCompare(b.engineer.name));
      return ranked[0]?.engineer ?? suggestedManagerEngineer;
    };

    const openManagerPaymentModal = (order: ServiceOrder) => {
      const due = Math.max(clientOrderDebt(order), 0);
      setManagerPaymentModal({
        orderId: order.id,
        amount: String(due || order.repairPrice || order.estimatedAmount || ''),
        due,
        method: order.repairPaymentMethod ?? 'наличные',
      });
    };

    const confirmManagerPaymentModal = () => {
      if (!managerPaymentModal) return;
      const targetOrder = managerSourceOrders.find((order) => order.id === managerPaymentModal.orderId);
      if (!targetOrder) {
        props.notifyUser('Замовлення не знайдено.');
        return;
      }
      const amountValue = Number(managerPaymentModal.amount);
      if (!Number.isFinite(amountValue) || amountValue <= 0) {
        props.notifyUser('Вкажіть суму оплати.');
        return;
      }
      const due = Math.max(clientOrderDebt(targetOrder), 0);
      const paymentKind: SimpleLedgerPaymentKind = amountValue + 0.01 >= due ? 'оплата' : 'частичная';
      const paymentId = props.acceptSimpleManagerPayment(targetOrder.id, String(amountValue), managerPaymentModal.method, paymentKind);
      if (!paymentId) return;
      pulseManagerOrder(targetOrder.id);
      setManagerPaymentModal(null);
      if (amountValue + 0.01 >= due) {
        setManagerPostPaymentOrderId(targetOrder.id);
      }
    };

    const openManagerIssueModal = (orderId: string) => {
      setManagerIssueModal({ orderId });
    };

    const confirmManagerIssueModal = () => {
      if (!managerIssueModal) return;
      runManagerStatusAction(managerIssueModal.orderId, 'Видано');
      pulseManagerOrder(managerIssueModal.orderId);
      setManagerIssueModal(null);
    };

    const managerSearchResults = managerVisibleOrders;
    const managerSearchExactOrder = findExactOrderMatch(managerSourceOrders, managerSearch);
    const managerSearchMatchedClient = managerSearchExactOrder ? null : findClientBySearch(props.customerList, managerSearch);
    const existingClientByPhone = props.quickPhone.trim()
      ? props.customerList.find((client) => extractDigits(client.phone) === extractDigits(props.quickPhone))
      : null;

    const openManagerOrderFromSearch = (order: ServiceOrder) => {
      setManagerFilter('all');
      openManagerOrderDetails(order.id);
      setManagerSearch(order.id);
      setShowManagerCreateForm(false);
      pulseManagerOrder(order.id);
    };

    const handleManagerSearchSubmit = () => {
      const needle = managerSearch.trim();
      if (!needle) return;
      const exact = findExactOrderMatch(managerSourceOrders, needle);
      if (showManagerCreateForm && (exact || managerSearchMatchedClient)) {
        return;
      }
      if (exact) {
        openManagerOrderFromSearch(exact);
        return;
      }
      if (managerSearchResults.length === 1) {
        openManagerOrderFromSearch(managerSearchResults[0]);
        return;
      }
      if (managerSearchResults.length === 0) {
        props.notifyUser('Нічого не знайдено.');
        return;
      }
      props.notifyUser(`Знайдено ${managerSearchResults.length} замовлень.`);
    };

    const managerActionMeta = (order: ServiceOrder) => {
      const remainingAmount = orderDebtAmount(order);
      const canMoveToRepair = order.status === 'Прийнято' && Boolean(order.engineer);
      const canMoveToReady = simpleRepairStatus(order.status) === 'В ремонті' && Boolean(order.engineerWorkCompletedAt || order.returnedToCellAt);
      const canTakePayment = !order.contractId && order.status !== 'Видано' && remainingAmount > 0 && ['Готовий до видачі', 'Очікує оплати'].includes(order.status);
      const actState = orderActStatus(order);
      const nextAction = getNextAction(order, actState.status);
      const blockReason = getBlockReason(order, actState.status);
      if (!order.engineer) {
        const nextEngineer = suggestManagerEngineer();
        return {
          signal: 'warning' as const,
          signalLabel: 'Без інженера',
          actionLabel: 'Призначити',
          action: nextEngineer ? () => props.patchSimpleManagerOrder(order.id, { engineer: nextEngineer.name }) : undefined,
          reason: 'Немає інженера',
        };
      }
      if (canMoveToRepair) {
        return {
          signal: 'neutral' as const,
          signalLabel: 'В роботу',
          actionLabel: 'В роботу',
          action: () => runManagerStatusAction(order.id, 'В ремонті'),
          reason: '',
        };
      }
      if (canMoveToReady) {
        return {
          signal: 'neutral' as const,
          signalLabel: 'Завершити',
          actionLabel: 'Завершити',
          action: () => runManagerStatusAction(order.id, 'Готово'),
          reason: '',
        };
      }
      if (canTakePayment) {
        return {
          signal: 'warning' as const,
          signalLabel: 'Не оплачено',
          actionLabel: nextAction === 'Оплатити' ? 'Оплатити' : 'Прийняти оплату',
          action: () => openManagerPaymentModal(order),
          reason: 'Є борг',
        };
      }
      if (nextAction === 'Підтвердити оплату') {
        const latestPendingPayment = order.payments.find((payment) => paymentNeedsConfirmation(payment));
        return {
          signal: 'warning' as const,
          signalLabel: 'Платіж не підтверджено',
          actionLabel: 'Підтвердити оплату',
          action: () => latestPendingPayment && props.confirmSimpleManagerPayment(latestPendingPayment.id),
          reason: blockReason ?? 'Платіж не підтверджено',
        };
      }
      if (nextAction === 'Роздрукувати акт' || nextAction === 'Підписати акт') {
        return {
          signal: actState.tone,
          signalLabel: actState.label,
          actionLabel: nextAction,
          action: () => actState.status === 'Роздруковано'
            ? props.signOrderAct(order)
            : (props.ensureOrderDocumentRecord('Акт надання послуг', order), props.logOrderDocumentPrint(order, 'Акт надання послуг'), props.printDocument('Акт надання послуг', 'service_order', order.id, order.client)),
          reason: blockReason ?? (actState.status === 'Роздруковано' ? 'Акт не підписано' : 'Акт не створено'),
        };
      }
      if (canRelease(order, actState.status)) {
        return {
          signal: 'success' as const,
          signalLabel: 'Можна видати',
          actionLabel: 'Видати',
          action: () => openManagerIssueModal(order.id),
          reason: '',
        };
      }
      if (remainingAmount > 0 && order.status === 'Видано') {
        return {
          signal: 'danger' as const,
          signalLabel: 'Борг',
          actionLabel: '',
          action: undefined,
          reason: 'Не закрито борг',
        };
      }
      if (simpleRepairStatus(order.status) === 'В ремонті' && !order.engineerWorkCompletedAt) {
        return {
          signal: 'neutral' as const,
          signalLabel: 'В роботі',
          actionLabel: '',
          action: undefined,
          reason: 'Інженер ще працює',
        };
      }
      if (blockReason === 'Є борг') {
        return {
          signal: 'danger' as const,
          signalLabel: 'Не можна видати',
          actionLabel: '',
          action: undefined,
          reason: 'Є борг',
        };
      }
      if (blockReason === 'Платіж не підтверджено') {
        return {
          signal: 'warning' as const,
          signalLabel: 'Платіж не підтверджено',
          actionLabel: '',
          action: undefined,
          reason: 'Платіж не підтверджено',
        };
      }
      return {
        signal: 'neutral' as const,
        signalLabel: 'В роботі',
        actionLabel: '',
        action: undefined,
        reason: 'Зараз без дії',
      };
    };

    const orderGroupPriority = (order: ServiceOrder) => {
      const remaining = clientOrderDebt(order);
      if (remaining > 0) return 1;
      if (!order.engineer) return 2;
      if (order.status === 'Готовий до видачі') return 3;
      return 4;
    };

    const statusAgeDays = (order: ServiceOrder) => Math.max(1, daysSince(order.statusChangedAt));
    const managerAgeTone = (order: ServiceOrder) => {
      const days = statusAgeDays(order);
      if (days >= 5) return 'danger';
      if (days >= 3) return 'warning';
      return 'neutral';
    };

    const managerGroupSort = (a: ServiceOrder, b: ServiceOrder) => {
      const priorityDiff = orderGroupPriority(a) - orderGroupPriority(b);
      if (priorityDiff !== 0) return priorityDiff;
      const ageDiff = statusAgeDays(b) - statusAgeDays(a);
      if (ageDiff !== 0) return ageDiff;
      return managerOrderPriorityScore(b) - managerOrderPriorityScore(a);
    };

    const searchRowSignalTone = (order: ServiceOrder) => {
      const remaining = clientOrderDebt(order);
      const ageDays = statusAgeDays(order);
      if (remaining > 0 || order.status === 'Видано') return 'danger';
      if (simpleRepairStatus(order.status) === 'Готово' || order.status === 'Готовий до видачі') return 'success';
      if (ageDays >= 3) return 'warning';
      return 'neutral';
    };

    const searchRowPriority = (order: ServiceOrder) => {
      const tone = searchRowSignalTone(order);
      if (tone === 'danger') return 1;
      if (tone === 'warning') return 2;
      if (tone === 'success') return 3;
      return 4;
    };

    const ageStatusLabel = (days: number) => {
      if (days === 1) return '1 день';
      if (days >= 2 && days <= 4) return `${days} дні`;
      return `${days} днів`;
    };

    const managerMatchedOrders = managerSearch.trim()
      ? [...managerSourceOrders]
        .filter((order) => !['Скасовано', 'Закрито'].includes(order.status))
        .filter((order) => matchesOrderSearch(order, managerSearch))
        .sort((a, b) => {
          const priorityDiff = searchRowPriority(a) - searchRowPriority(b);
          if (priorityDiff !== 0) return priorityDiff;
          return managerGroupSort(a, b);
        })
        .slice(0, 8)
      : [];
    const managerPostPaymentOrder = managerPostPaymentOrderId ? managerSourceOrders.find((order) => order.id === managerPostPaymentOrderId) ?? props.allOrders.find((order) => order.id === managerPostPaymentOrderId) : null;

    const canTakePaymentFromSearch = (order: ServiceOrder) => {
      const totalAmount = Math.max(order.repairPrice ?? order.estimatedAmount ?? orderTotals(order).total, 0);
      const paidAmount = order.payments.filter(paymentCountsAsApplied).reduce((sum, payment) => sum + payment.amount, 0);
      const remainingAmount = Math.max(totalAmount - paidAmount, 0);
      return !order.contractId && order.status !== 'Видано' && remainingAmount > 0 && ['Готовий до видачі', 'Очікує оплати'].includes(order.status);
    };

    const orderActStatus = (order: ServiceOrder) => {
      const actStatus = getOrderActStatus(props.documents, order.id);
      if (actStatus === 'Не створено') {
        return {
          status: actStatus,
          tone: 'danger' as const,
          icon: '🔴',
          label: 'Акт не створено',
        };
      }
      if (actStatus === 'Підписано') {
        return {
          status: actStatus,
          tone: 'success' as const,
          icon: '🟢',
          label: 'Акт підписано',
        };
      }
      if (actStatus === 'Роздруковано') {
        return {
          status: actStatus,
          tone: 'warning' as const,
          icon: '🟡',
          label: 'Акт роздруковано',
        };
      }
      return {
        status: actStatus,
        tone: 'danger' as const,
        icon: '🔴',
        label: 'Акт не створено',
      };
    };

    const canIssueFromSearch = (order: ServiceOrder) => {
      return canRelease(order, getOrderActStatus(props.documents, order.id));
    };

    const searchRowIssueGuard = (order: ServiceOrder) => {
      const remainingAmount = orderDebtAmount(order);
      const latestPendingPayment = order.payments.find((payment) => paymentNeedsConfirmation(payment));
      const actState = orderActStatus(order);
      const nextAction = getNextAction(order, actState.status);
      const releaseBlockReason = getBlockReason(order, actState.status);
      if (canIssueFromSearch(order)) {
        return { canIssue: true as const };
      }
      if (nextAction === 'Підтвердити оплату' && latestPendingPayment) {
        return {
          canIssue: false as const,
          reason: releaseBlockReason ?? pendingPaymentReason(order),
          actionLabel: 'Підтвердити платіж',
          action: () => props.confirmSimpleManagerPayment(latestPendingPayment.id),
        };
      }
      if (nextAction === 'Оплатити' && remainingAmount > 0) {
        return {
          canIssue: false as const,
          reason: `Є борг: ${money(remainingAmount)}`,
          actionLabel: 'Оплатити',
          action: () => openManagerPaymentModal(order),
        };
      }
      if (order.status === 'Видано') {
        return {
          canIssue: false as const,
          reason: 'Замовлення вже видано',
          actionLabel: 'Перейти',
          action: () => openManagerOrderFromSearch(order),
        };
      }
      if (nextAction === 'Роздрукувати акт' || nextAction === 'Підписати акт') {
        return {
          canIssue: false as const,
          reason: actState.status === 'Роздруковано' ? 'Акт не підписано' : 'Акт не створено',
          actionLabel: actState.status === 'Роздруковано' ? 'Підписати акт' : 'Роздрукувати акт',
          action: () => actState.status === 'Роздруковано'
            ? props.signOrderAct(order)
            : (props.ensureOrderDocumentRecord('Акт надання послуг', order), props.logOrderDocumentPrint(order, 'Акт надання послуг'), props.printDocument('Акт надання послуг', 'service_order', order.id, order.client)),
        };
      }
      if (nextAction === 'Перейти до ремонту') {
        return {
          canIssue: false as const,
          reason: releaseBlockReason ?? 'Замовлення не завершено',
          actionLabel: 'Перейти до ремонту',
          action: () => openManagerOrderFromSearch(order),
        };
      }
      return {
        canIssue: false as const,
        reason: releaseBlockReason ?? 'Замовлення ще не готове до видачі',
        actionLabel: 'Перейти до ремонту',
        action: () => openManagerOrderFromSearch(order),
      };
    };

    const searchRowMainAction = (order: ServiceOrder) => {
      const issueGuard = searchRowIssueGuard(order);
      const actState = orderActStatus(order);
      const nextAction = getNextAction(order, actState.status);
      if (nextAction === 'Оплатити') {
        return {
          label: 'Оплатити',
          run: () => openManagerPaymentModal(order),
        };
      }
      if (nextAction === 'Підтвердити оплату') {
        const latestPendingPayment = order.payments.find((payment) => paymentNeedsConfirmation(payment));
        return {
          label: 'Підтвердити оплату',
          run: () => latestPendingPayment && props.confirmSimpleManagerPayment(latestPendingPayment.id),
        };
      }
      if (nextAction === 'Роздрукувати акт' || nextAction === 'Підписати акт') {
        return {
          label: actState.status === 'Роздруковано' ? 'Підписати акт' : 'Роздрукувати акт',
          run: () => actState.status === 'Роздруковано'
            ? props.signOrderAct(order)
            : (props.ensureOrderDocumentRecord('Акт надання послуг', order), props.logOrderDocumentPrint(order, 'Акт надання послуг'), props.printDocument('Акт надання послуг', 'service_order', order.id, order.client)),
        };
      }
      if (issueGuard.canIssue) {
        return {
          label: 'Видати',
          run: () => openManagerIssueModal(order.id),
        };
      }
      return {
        label: 'Перейти до ремонту',
        run: () => openManagerOrderFromSearch(order),
      };
    };

    const groupedManagerOrders = {
      urgent: managerVisibleOrders.filter((order) => {
        const remaining = clientOrderDebt(order);
        return remaining > 0 && order.status === 'Видано'
          || (simpleRepairStatus(order.status) === 'Готово' && remaining > 0)
          || (simpleRepairStatus(order.status) === 'Готово' && order.status !== 'Видано' && daysSince(order.statusChangedAt) >= 1)
          || (!order.engineer && hoursSince(order.statusChangedAt) >= 24)
          || hoursSince(order.statusChangedAt) >= 72;
      }).sort(managerGroupSort),
      attention: managerVisibleOrders.filter((order) => {
        const remaining = clientOrderDebt(order);
        const isUrgent = remaining > 0 && order.status === 'Видано'
          || (simpleRepairStatus(order.status) === 'Готово' && remaining > 0)
          || (!order.engineer && hoursSince(order.statusChangedAt) >= 24)
          || hoursSince(order.statusChangedAt) >= 72;
        if (isUrgent) return false;
        return !order.clientNotified
          || remaining > 0
          || !order.engineer
          || order.payments.some((payment) => paymentNeedsConfirmation(payment));
      }).sort(managerGroupSort),
      normal: managerVisibleOrders.filter((order) => {
        const remaining = clientOrderDebt(order);
        const isUrgent = remaining > 0 && order.status === 'Видано'
          || (simpleRepairStatus(order.status) === 'Готово' && remaining > 0)
          || (!order.engineer && hoursSince(order.statusChangedAt) >= 24)
          || hoursSince(order.statusChangedAt) >= 72;
        const isAttention = !order.clientNotified
          || remaining > 0
          || !order.engineer
          || order.payments.some((payment) => paymentNeedsConfirmation(payment));
        return !isUrgent && !isAttention;
      }).sort(managerGroupSort),
    };

    return (
      <div className="page-grid manager-orders-page">
        <section className="panel manager-orders-toolbar">
          <div className="manager-orders-toolbar-search">
            <div className="manager-orders-search-input">
              <input
                value={managerSearch}
                onChange={(event) => setManagerSearch(event.target.value)}
                onKeyDown={(event) => {
                  if (event.key === 'Enter') {
                    event.preventDefault();
                    handleManagerSearchSubmit();
                  }
                }}
                placeholder="Пошук: номер / телефон / клієнт / пристрій"
              />
              <button type="button" className="manager-orders-search-button" onClick={handleManagerSearchSubmit} aria-label="Запустити пошук">
                <Search size={16} />
              </button>
            </div>
            <button type="button" className="submit-button" onClick={openManagerCreateMode}>
              ➕ Нове замовлення
            </button>
          </div>
          {!showManagerCreateForm && (
            <div className="manager-orders-toolbar-filter">
              <button type="button" className={managerFilter === 'all' ? 'primary' : ''} onClick={() => setManagerFilter('all')}>Усі</button>
              <button type="button" className={managerFilter === 'Прийнято' ? 'primary' : ''} onClick={() => setManagerFilter('Прийнято')}>Прийнято</button>
              <button type="button" className={managerFilter === 'В ремонті' ? 'primary' : ''} onClick={() => setManagerFilter('В ремонті')}>В ремонті</button>
              <button type="button" className={managerFilter === 'Готово' ? 'primary' : ''} onClick={() => setManagerFilter('Готово')}>Готово</button>
              <button type="button" className={managerFilter === 'Видано' ? 'primary' : ''} onClick={() => setManagerFilter('Видано')}>Видано</button>
              <button type="button" className={managerFilter === 'Борг' ? 'primary' : ''} onClick={() => setManagerFilter('Борг')}>Борг</button>
              <button type="button" className={managerFilter === 'Очікує оплату' ? 'primary' : ''} onClick={() => setManagerFilter('Очікує оплату')}>Очікує оплату</button>
            </div>
          )}
        </section>

        {!showManagerCreateForm && managerSearch.trim() && (
          <section className="panel manager-orders-search-assist">
            <div className="panel-heading">
              <h2>Результати пошуку</h2>
              <span>Швидкий доступ до замовлень і клієнтів</span>
            </div>
            <div className="manager-search-results">
              {managerMatchedOrders.map((order) => {
                const totalAmount = Math.max(order.repairPrice ?? order.estimatedAmount ?? orderTotals(order).total, 0);
                const paidAmount = order.payments.filter(paymentCountsAsApplied).reduce((sum, payment) => sum + payment.amount, 0);
                const remainingAmount = Math.max(totalAmount - paidAmount, 0);
                const ageDays = statusAgeDays(order);
                const ageTone = managerAgeTone(order);
                const signalTone = searchRowSignalTone(order);
                const actState = orderActStatus(order);
                const financeLabel = remainingAmount > 0 ? `Борг: ${money(remainingAmount)}` : `${money(paidAmount)} / ${money(totalAmount)}`;
                const mainAction = searchRowMainAction(order);
                const issueGuard = searchRowIssueGuard(order);
                return (
                  <div
                    key={order.id}
                    className={`manager-search-result-row manager-search-result-${signalTone}`}
                    onClick={() => openManagerOrderFromSearch(order)}
                    role="button"
                    tabIndex={0}
                    onKeyDown={(event) => {
                      if (event.key === 'Enter' || event.key === ' ') {
                        event.preventDefault();
                        openManagerOrderFromSearch(order);
                      }
                    }}
                  >
                    <div className="manager-search-result-main">
                      <strong>[{order.id.replace('ЗН-', 'ZN-')}] {order.client}</strong>
                      <span>{order.device} | {managerOrderStatusLabel(order.status)} | {order.engineer || 'Без інженера'}</span>
                      <span className={`manager-search-result-act manager-search-result-act-${actState.tone}`}>{actState.icon} {actState.label}</span>
                    </div>
                    <div className="manager-search-result-meta">
                      <span>{financeLabel}</span>
                      <span className={`manager-search-result-age manager-search-result-age-${ageTone}`}>{ageStatusLabel(ageDays)}</span>
                    </div>
                    <div className="manager-search-result-actions">
                      {!issueGuard.canIssue && (
                        <div className="manager-search-result-blocked">
                          <strong>НЕ МОЖНА ВИДАТИ</strong>
                          <span>{issueGuard.reason}</span>
                        </div>
                      )}
                      <button
                        type="button"
                        className="manager-search-result-primary"
                        onClick={(event) => {
                          event.stopPropagation();
                          (!issueGuard.canIssue && issueGuard.actionLabel ? issueGuard.action : mainAction.run)?.();
                        }}
                      >
                        {!issueGuard.canIssue && issueGuard.actionLabel ? issueGuard.actionLabel : mainAction.label}
                      </button>
                    </div>
                  </div>
                );
              })}
              {managerSearchMatchedClient && (
                <div className="manager-search-result-row manager-search-result-client">
                  <div className="manager-search-result-main">
                    <strong>{managerSearchMatchedClient.name}</strong>
                    <span>{managerSearchMatchedClient.phone}{managerSearchMatchedClient.taxId ? ` | ${managerSearchMatchedClient.taxId}` : ''}</span>
                  </div>
                  <div className="manager-search-result-actions">
                    <button
                      type="button"
                      onClick={() => {
                        setShowManagerCreateForm(true);
                        props.setQuickClientName(managerSearchMatchedClient.name);
                        props.setQuickPhone(managerSearchMatchedClient.phone);
                      }}
                    >
                      Створити замовлення
                    </button>
                    <button type="button" onClick={() => props.openClientRecord(managerSearchMatchedClient.phone, managerSearchMatchedClient.name)}>Відкрити клієнта</button>
                  </div>
                </div>
              )}
              {managerMatchedOrders.length === 0 && !managerSearchMatchedClient && (
                <div className="empty-state">Нічого не знайдено</div>
              )}
            </div>
          </section>
        )}

        {!showManagerCreateForm && managerPostPaymentOrder && (
          <section className="panel manager-orders-documents-assist">
            <div className="panel-heading">
              <h2>Оплату підтверджено</h2>
              <span>{managerPostPaymentOrder.id} · борг закрито, система веде далі</span>
            </div>
            {(() => {
              const actDocument = props.documents.find((document) => document.entityType === 'service_order' && document.entityId === managerPostPaymentOrder.id && document.kind === 'Акт надання послуг');
              const invoiceDocument = props.documents.find((document) => document.entityType === 'service_order' && document.entityId === managerPostPaymentOrder.id && document.kind === 'Рахунок на оплату');
              const deliveryDocument = props.documents.find((document) => document.entityType === 'service_order' && document.entityId === managerPostPaymentOrder.id && document.kind === 'Видаткова накладна');
              const taxInvoice = [...props.taxInvoices]
                .filter((item) => item.orderId === managerPostPaymentOrder.id)
                .sort((a, b) => (parseDateTime(b.createdAt)?.getTime() ?? 0) - (parseDateTime(a.createdAt)?.getTime() ?? 0))[0];
              return (
                <>
            <div className="manager-order-hints">
              <span className="manager-order-hint manager-order-hint-success">Замовлення переведено в "Готовий до видачі"</span>
              {(managerPostPaymentOrder.legalEntity || managerPostPaymentOrder.vatStatus) && (
                <span className="manager-order-hint manager-order-hint-warning">
                  {taxInvoice ? `${taxInvoice.number} · ${taxInvoice.status}` : (managerPostPaymentOrder.vatStatus || 'Очікує ПН')}
                </span>
              )}
            </div>
            <div className="manager-order-bottom-meta">
              <div className="manager-order-field"><strong>Акт</strong><div>{actDocument ? actDocument.status : 'Не створено'}</div></div>
              <div className="manager-order-field"><strong>Рахунок</strong><div>{invoiceDocument ? invoiceDocument.status : 'Не створено'}</div></div>
              <div className="manager-order-field"><strong>Видаткова</strong><div>{managerPostPaymentOrder.parts.length > 0 ? (deliveryDocument ? deliveryDocument.status : 'Не створено') : 'Не потрібна'}</div></div>
              {(managerPostPaymentOrder.legalEntity || managerPostPaymentOrder.vatStatus) && <div className="manager-order-field"><strong>ПН</strong><div>{taxInvoice ? `${taxInvoice.number} · ${taxInvoice.status}` : 'Не створено'}</div></div>}
            </div>
            <div className="manager-post-payment-actions">
              {!invoiceDocument && <button type="button" onClick={() => props.createServiceOrderDocument('Рахунок на оплату', managerPostPaymentOrder)}>Створити рахунок</button>}
              {invoiceDocument && <button type="button" onClick={() => props.printDocument('Рахунок на оплату', 'service_order', managerPostPaymentOrder.id, managerPostPaymentOrder.client)}>Роздрукувати рахунок</button>}
              {!actDocument && <button type="button" onClick={() => props.createServiceOrderDocument('Акт надання послуг', managerPostPaymentOrder)}>Створити акт</button>}
              {actDocument && <button type="button" onClick={() => props.printDocument('Акт надання послуг', 'service_order', managerPostPaymentOrder.id, managerPostPaymentOrder.client)}>Роздрукувати акт</button>}
              {invoiceDocument && (
                <>
                  <button type="button" onClick={() => { props.notifyUser('У діалозі друку оберіть "Зберегти як PDF".'); props.printDocument('Рахунок на оплату', 'service_order', managerPostPaymentOrder.id, managerPostPaymentOrder.client); }}>Скачати PDF</button>
                  <button type="button" onClick={() => props.printDocument('Рахунок на оплату', 'service_order', managerPostPaymentOrder.id, managerPostPaymentOrder.client)}>Повторно роздрукувати</button>
                </>
              )}
              {managerPostPaymentOrder.parts.length > 0 && !deliveryDocument && (
                <button type="button" onClick={() => props.createServiceOrderDocument('Видаткова накладна', managerPostPaymentOrder)}>Створити видаткову накладну</button>
              )}
              {managerPostPaymentOrder.parts.length > 0 && deliveryDocument && (
                <button type="button" onClick={() => props.printDocument('Видаткова накладна', 'service_order', managerPostPaymentOrder.id, managerPostPaymentOrder.client)}>Роздрукувати видаткову накладну</button>
              )}
              {actDocument && actDocument.status !== 'Підписано' && (
                <button type="button" onClick={() => props.signOrderAct(managerPostPaymentOrder)}>Підписати акт</button>
              )}
              {(managerPostPaymentOrder.legalEntity || managerPostPaymentOrder.vatStatus) && (
                taxInvoice
                  ? taxInvoice.status === 'Зареєстровано'
                    ? <button type="button" disabled>ПН зареєстрована</button>
                    : <button type="button" onClick={() => props.registerTaxInvoice(managerPostPaymentOrder.id)}>Позначити як зареєстровано</button>
                  : <button type="button" onClick={() => props.createTaxInvoiceForOrder(managerPostPaymentOrder.id)}>Створити ПН</button>
              )}
              <button type="button" className="submit-button" onClick={() => { setManagerActiveOrderId(managerPostPaymentOrder.id); setManagerPostPaymentOrderId(''); }}>До видачі</button>
            </div>
                </>
              );
            })()}
          </section>
        )}

        {!showManagerCreateForm && (
          <section className="manager-orders-signals">
            {managerAlerts.debt > 0 && <span className="manager-order-hint manager-order-hint-danger">Борг: {managerAlerts.debt}</span>}
            {managerAlerts.unpaid > 0 && <span className="manager-order-hint manager-order-hint-warning">Не оплачено: {managerAlerts.unpaid}</span>}
            {managerAlerts.readyWaiting > 0 && <span className="manager-order-hint manager-order-hint-danger">Готово, не видано: {managerAlerts.readyWaiting}</span>}
            {managerAlerts.noEngineer > 0 && <span className="manager-order-hint manager-order-hint-warning">Без інженера: {managerAlerts.noEngineer}</span>}
            {managerAlerts.debt === 0 && managerAlerts.unpaid === 0 && managerAlerts.readyWaiting === 0 && managerAlerts.noEngineer === 0 && (
              <span className="manager-order-hint manager-order-hint-neutral">Все під контролем</span>
            )}
          </section>
        )}

        {showManagerCreateForm ? (
          <section className="panel manager-orders-create">
            <div className="panel-heading">
              <h2>Нове замовлення</h2>
              <span>Заповніть клієнта, телефон, пристрій, проблему та оберіть договір або режим без договору.</span>
            </div>
            <div className="table">
              <label>
                Імʼя / ПІБ або назва фірми
                <input value={props.quickClientName} onChange={(event) => props.setQuickClientName(event.target.value)} placeholder="Імʼя клієнта або назва фірми" />
              </label>
              <label>
                Телефон
                <input value={props.quickPhone} onChange={(event) => props.setQuickPhone(event.target.value)} placeholder="+380..." />
              </label>
              {existingClientByPhone && (
                <div className="manager-order-hints">
                  <span className="manager-order-hint manager-order-hint-warning">Клієнт з таким телефоном вже існує</span>
                  <button type="button" onClick={() => props.openClientRecord(existingClientByPhone.phone, existingClientByPhone.name)}>Відкрити клієнта</button>
                  <button
                    type="button"
                    onClick={() => {
                      props.setQuickClientName(existingClientByPhone.name);
                      props.setQuickPhone(existingClientByPhone.phone);
                    }}
                  >
                    Створити замовлення цьому клієнту
                  </button>
                </div>
              )}
              <label>
                Пристрій
                <input value={props.quickDevice} onChange={(event) => props.setQuickDevice(event.target.value)} placeholder="HP LaserJet" />
              </label>
              <label>
                Проблема
                <input value={props.quickProblem} onChange={(event) => props.setQuickProblem(event.target.value)} placeholder="Не друкує" />
              </label>
              <label>
                Договір
                <select value={props.quickContractId} onChange={(event) => props.setQuickContractId(event.target.value)}>
                  <option value="">Без договору</option>
                  {managerContracts.map((contract) => <option key={contract.id} value={contract.id}>{contract.id} · {contract.client}</option>)}
                </select>
              </label>
            </div>
            {props.quickClientDebtWarning && (
              <div className="manager-order-hints">
                <span className="manager-order-hint manager-order-hint-danger">{props.quickClientDebtWarning}</span>
              </div>
            )}
            <div className="action-row">
              <button type="button" className="submit-button" onClick={submitManagerOrder}>Зберегти замовлення</button>
              <button type="button" onClick={cancelManagerCreateMode}>Скасувати / Назад до замовлень</button>
            </div>
          </section>
        ) : (
        <>
        <div className="manager-orders-workspace">
          <section className="panel manager-orders-list manager-orders-list-full">
            <div className="panel-heading">
              <h2>{props.allRoleOrders ? 'Мої замовлення' : 'Всі замовлення'}</h2>
              <span>{managerVisibleOrders.length}</span>
            </div>
            <div className="manager-orders-list-body">
              {managerVisibleOrders.map((order) => {
                const debtSnapshot = orderDebtSnapshot(order);
                const remainingForBadge = debtSnapshot.remainingDebt;
                const strictStatusLabel = strictManagerWorkflowStatus(order);
                const actionMeta = managerActionMeta(order);
                const canQuickPay = remainingForBadge > 0 && order.status !== 'Видано';
                const canQuickIssue = simpleRepairStatus(order.status) === 'Готово' && order.status !== 'Видано';
                return (
                  <button
                    type="button"
                    key={order.id}
                    className={`manager-order-list-row manager-order-compact-row${managerActiveOrderId === order.id && isManagerOrderDetailOpen ? ' is-active' : ''}${order.status === 'Видано' ? ' is-issued' : ''}${actionMeta.signal === 'danger' ? ' is-problem' : actionMeta.signal === 'warning' ? ' is-warning' : actionMeta.signal === 'success' ? ' is-success' : ''}`}
                    onClick={() => openManagerOrderDetails(order.id)}
                  >
                    <span className="manager-order-primary">
                      <span className="manager-order-id-badge">{order.id}</span>
                      <span className="manager-order-device-icon" aria-hidden="true">{deviceIcon(order.device)}</span>
                      <span className="manager-order-primary-copy">
                        <span className="manager-order-device-text">{order.device}</span>
                        <span className="manager-order-client-text">{order.client}</span>
                      </span>
                    </span>
                    <span className="manager-order-secondary">
                      <span className="manager-order-cell manager-order-cell-status">
                        <span className={`manager-order-status-badge ${simpleRepairStatusClass(simpleRepairStatus(order.status))}`}>{strictStatusLabel}</span>
                      </span>
                      <span className="manager-order-metric">
                        <span className="manager-order-metric-label">Сума</span>
                        <span className="manager-order-cell manager-order-cell-money">{money(debtSnapshot.total)}</span>
                      </span>
                      <span className="manager-order-metric">
                        <span className="manager-order-metric-label">Борг</span>
                        <span className={`manager-order-cell manager-order-cell-money ${remainingForBadge > 0 ? 'manager-order-cell-debt manager-order-debt-badge' : ''}`}>{remainingForBadge > 0 ? `● ${money(remainingForBadge)}` : '—'}</span>
                      </span>
                      <span className="manager-order-metric manager-order-metric-date">
                        <span className="manager-order-metric-label">Дата</span>
                        <span className="manager-order-cell manager-order-cell-date">{order.intakeDate}</span>
                      </span>
                      {(canQuickPay || canQuickIssue) && (
                        <span className="manager-order-inline-actions">
                          {canQuickPay && (
                            <button
                              type="button"
                              className="manager-order-inline-action manager-order-inline-action-payment"
                              onClick={(event) => {
                                event.stopPropagation();
                                openManagerPaymentModal(order);
                              }}
                            >
                              Оплата
                            </button>
                          )}
                          {canQuickIssue && (
                            <button
                              type="button"
                              className="manager-order-inline-action manager-order-inline-action-issue"
                              onClick={(event) => {
                                event.stopPropagation();
                                openManagerIssueModal(order.id);
                              }}
                            >
                              Видати
                            </button>
                          )}
                        </span>
                      )}
                    </span>
                  </button>
                );
              })}
              {managerVisibleOrders.length === 0 && <div className="empty-state">Немає замовлень</div>}
            </div>
          </section>
        </div>
        {isManagerOrderDetailOpen && activeManagerOrder && (
          <div className="manager-order-modal-backdrop" onClick={() => setIsManagerOrderDetailOpen(false)}>
            <section className="panel manager-order-focus manager-order-modal" onClick={(event) => event.stopPropagation()}>
              <div className="panel-heading">
                <h2>Замовлення {activeManagerOrder.id}</h2>
                <button type="button" onClick={() => setIsManagerOrderDetailOpen(false)}>Закрити</button>
              </div>
            {(() => {
              const order = activeManagerOrder;
              const linkedContract = order.contractId ? props.contracts.find((contract) => contract.id === order.contractId) : undefined;
              const linkedAct = order.contractActId ? props.contractActs.find((act) => act.id === order.contractActId) : undefined;
              const debtSnapshot = orderDebtSnapshot(order);
              const clientDebtTotal = props.allOrders
                .filter((entry) => entry.client.trim().toLowerCase() === order.client.trim().toLowerCase() || entry.phone.replace(/\D/g, '') === order.phone.replace(/\D/g, ''))
                .reduce((sum, entry) => sum + clientOrderDebt(entry), 0);
              const strictStatusLabel = strictManagerWorkflowStatus(order);
              const orderCost = order.parts.reduce((sum, part) => sum + part.qty * part.cost, 0);
              const totalAmount = Math.max(order.repairPrice ?? order.estimatedAmount ?? orderTotals(order).total, 0);
              const paidAmount = order.payments.filter(paymentCountsAsApplied).reduce((sum, payment) => sum + payment.amount, 0);
              const remainingAmount = orderDebtAmount(order);
              const pendingPayments = order.payments.filter((payment) => paymentNeedsConfirmation(payment));
              const actState = orderActStatus(order);
              const pendingReason = pendingPaymentReason(order);
              const releaseBlockReason = getBlockReason(order, actState.status);
              const primaryAction = getNextAction(order, actState.status);
              const primaryActionHint = getNextActionHint(order, actState.status);
              const paymentDraft = managerPaymentDraft(order);
              const paymentVisualStatus = simpleLedgerPaymentLabel(order);
              const filteredProducts = managerAvailableProducts.filter((product) => product.name.toLowerCase().includes(partSearch.trim().toLowerCase()));
              const canMoveToRepair = order.status === 'Прийнято' && Boolean(order.engineer);
              const canMoveToReady = simpleRepairStatus(order.status) === 'В ремонті' && Boolean(order.engineerWorkCompletedAt || order.returnedToCellAt);
              const canTakePayment = !linkedContract && order.status !== 'Видано' && remainingAmount > 0 && ['Готовий до видачі', 'Очікує оплати'].includes(order.status);
              const canAccountToContract = Boolean(linkedContract) && !order.contractAccountedAt && ['Готовий до видачі', 'Очікує оплати'].includes(order.status);
              const canCancelOrder = !['Скасовано', 'Видано', 'Закрито'].includes(order.status);
              const canReturnToWork = ['Готовий до видачі', 'Очікує оплати'].includes(order.status) || strictStatusLabel === 'Готово до видачі';
              const canRefundMoney = paidAmount > 0;
              const can_release = canRelease(order, actState.status);
              const issuedWithDebt = order.status === 'Видано' && remainingAmount > 0;
              const canIssueLabel = issuedWithDebt ? 'ВИДАНО З БОРГОМ' : can_release ? 'МОЖНА ВИДАТИ' : 'НЕ МОЖНА ВИДАТИ';
              const canIssueToneClass = issuedWithDebt ? 'manager-order-issue-debt' : can_release ? 'manager-order-issue-allowed' : 'manager-order-issue-blocked';
              const latestPayment = order.payments[0];
              const lastPaymentMethod = latestPayment?.method === 'Готівка' ? 'Готівка' : latestPayment?.method === 'Картка' ? 'Термінал' : latestPayment?.method === 'Безготівка' ? 'Банк' : '—';
              const lastPaymentDate = latestPayment?.date ?? '—';
              const lastPaymentStatus = latestPayment ? paymentProcessingLabel(latestPayment) : '—';
              const orderProfit = totalAmount - orderCost;
              const nextManagerAction = !order.engineer
                ? { label: 'Призначити інженера', run: () => undefined }
                : canMoveToRepair
                  ? { label: 'В роботу', run: () => runManagerStatusAction(order.id, 'В ремонті') }
                  : canMoveToReady
                    ? { label: 'Завершити ремонт', run: () => runManagerStatusAction(order.id, 'Готово') }
                    : primaryAction === 'Оплатити' && canTakePayment
                      ? { label: 'Оплатити', run: () => openManagerPaymentModal(order) }
                      : primaryAction === 'Підтвердити оплату'
                        ? {
                            label: 'Підтвердити оплату',
                            run: () => {
                              const latestPendingPayment = order.payments.find((payment) => paymentNeedsConfirmation(payment));
                              if (latestPendingPayment) props.confirmSimpleManagerPayment(latestPendingPayment.id);
                            },
                          }
                        : primaryAction === 'Роздрукувати акт' || primaryAction === 'Підписати акт'
                          ? {
                              label: primaryAction,
                              run: () => actState.status === 'Роздруковано'
                                ? props.signOrderAct(order)
                                : (props.ensureOrderDocumentRecord('Акт надання послуг', order), props.logOrderDocumentPrint(order, 'Акт надання послуг'), props.printDocument('Акт надання послуг', 'service_order', order.id, order.client)),
                            }
                          : can_release
                            ? { label: 'Видати', run: () => openManagerIssueModal(order.id) }
                            : primaryAction === 'Перейти до ремонту'
                              ? { label: 'Перейти до ремонту', run: () => setManagerActiveOrderId(order.id) }
                              : null;
              const nextManagerReason = !order.engineer
                ? 'Немає інженера'
                : order.status === 'Прийнято'
                  ? 'Замовлення ще не взято в роботу'
                  : simpleRepairStatus(order.status) === 'В ремонті' && !order.engineerWorkCompletedAt
                    ? 'Інженер ще не завершив роботу'
                    : simpleRepairStatus(order.status) === 'В ремонті' && order.engineerWorkCompletedAt && !order.returnedToCellAt
                      ? 'Потрібно повернути замовлення в комірку'
                    : releaseBlockReason === 'Платіж не підтверджено'
                      ? (pendingReason || 'Платіж не підтверджено')
                      : releaseBlockReason === 'Акт не підписано'
                        ? (actState.status === 'Роздруковано' ? 'Акт не підписано' : 'Акт не створено')
                        : releaseBlockReason === 'Є борг'
                          ? 'Є борг або очікується оплата'
                          : can_release
                            ? 'Можна рухати далі'
                            : releaseBlockReason === 'Замовлення не завершено'
                              ? 'Замовлення не завершено'
                              : primaryActionHint;
              const notifyEvent: NotificationEvent = order.status === 'Очікує оплати'
                ? 'Очікує оплату'
                : order.status === 'Готовий до видачі'
                  ? 'Готово до видачі'
                  : simpleRepairStatus(order.status) === 'В ремонті'
                    ? 'Ремонт розпочато'
                    : 'Нагадування';
              return (
                <div className={`manager-order-card manager-order-focus-card${managerPulseOrderId === order.id ? ' is-pulse' : ''}`}>
                  <div className="manager-order-card-header">
                    <div className="manager-order-card-title">
                      <div className="panel-heading manager-order-heading-row">
                        <h2>{order.id}</h2>
                        <span className={`manager-order-status-badge ${simpleRepairStatusClass(simpleRepairStatus(order.status))}`}>{strictStatusLabel}</span>
                      </div>
                      <div className="manager-order-hints">
                        {remainingAmount > 0 && <span className="manager-order-hint manager-order-hint-danger">Борг</span>}
                        {!order.engineer && <span className="manager-order-hint manager-order-hint-warning">Без інженера</span>}
                        {clientDebtTotal > 0 && <span className="manager-order-hint manager-order-hint-danger">Клієнт винен {money(clientDebtTotal)}</span>}
                      </div>
                    </div>
                  </div>

                  <div className="manager-order-focus-layout">
                    <div className="manager-order-focus-main">
                      <div className="manager-order-topline manager-order-section">
                        <div className={`manager-order-issue-flag ${canIssueToneClass}`}>
                          <strong>{canIssueLabel}</strong>
                          <span>{nextManagerReason}</span>
                        </div>
                        <div className="manager-order-topline-money">
                          <div><span>Сума</span><strong>{money(totalAmount)}</strong></div>
                          <div><span>Оплачено</span><strong>{money(paidAmount)}</strong></div>
                          <div><span>Залишок</span><strong>{money(remainingAmount)}</strong></div>
                        </div>
                        <div className="manager-order-inline-actions">
                          {nextManagerAction ? (
                            <button type="button" className="submit-button" onClick={nextManagerAction.run}>{nextManagerAction.label}</button>
                          ) : (
                            <button type="button" disabled>{nextManagerReason}</button>
                          )}
                        </div>
                      </div>

                      <div className="manager-order-workline manager-order-section">
                        <div className="manager-order-field">
                          <strong>Клієнт</strong>
                          <div>{order.client}</div>
                        </div>
                        <div className="manager-order-field">
                          <strong>Телефон</strong>
                          <div><a href={`tel:${order.phone}`}>{order.phone}</a></div>
                        </div>
                        <div className="manager-order-field">
                          <strong>Пристрій</strong>
                          <div>{order.device}</div>
                        </div>
                        <div className="manager-order-field">
                          <strong>Проблема</strong>
                          <div>{order.issue}</div>
                        </div>
                        <label className="manager-order-field">
                          <span>Інженер</span>
                          <select
                            value={props.users.find((user) => user.role === 'Інженер' && user.name === order.engineer)?.id ?? ''}
                            onKeyDown={handleManagerFieldKeyDown}
                            onChange={(event) => {
                              const nextEngineer = props.users.find((user) => user.id === event.target.value && user.role === 'Інженер');
                              if (nextEngineer) props.patchSimpleManagerOrder(order.id, { engineer: nextEngineer.name });
                            }}
                          >
                            <option value="">Оберіть інженера</option>
                            {managerEngineers.map((engineer) => <option key={engineer.id} value={engineer.id}>{engineer.name}</option>)}
                          </select>
                        </label>
                        <div className="manager-order-field">
                          <strong>Дата</strong>
                          <div>{order.intakeDate}</div>
                        </div>
                        <label className="manager-order-field manager-order-focus-comment">
                          <span>Коментар менеджера</span>
                          <textarea
                            value={order.intakeComment ?? ''}
                            rows={2}
                            onKeyDown={handleManagerFieldKeyDown}
                            onChange={(event) => props.patchSimpleManagerOrder(order.id, { intakeComment: event.target.value })}
                          />
                        </label>
                      </div>
                    </div>

                    <div className="manager-order-focus-side">
                      <div className="manager-order-payment manager-order-section">
                        <div className="panel-heading">
                          <h2>Фінанси</h2>
                          {(canTakePayment || canAccountToContract) && (
                            <button
                              type="button"
                              className="submit-button"
                              onClick={() => {
                                if (linkedContract) {
                                  props.accountSimpleManagerOrderToContract(order.id);
                                } else {
                                  openManagerPaymentModal(order);
                                }
                                pulseManagerOrder(order.id);
                              }}
                            >
                              {linkedContract ? 'Учесть в договор' : 'Прийняти оплату'}
                            </button>
                          )}
                        </div>
                        <div className="manager-order-payment-grid">
                          <label className="manager-order-money-primary">Сума ремонту<input type="number" min={0} value={order.repairPrice ?? ''} onKeyDown={handleManagerFieldKeyDown} onChange={(event) => props.patchSimpleManagerOrder(order.id, { repairPrice: event.target.value ? Number(event.target.value) : undefined })} /></label>
                          <label>Оплачено<input value={money(paidAmount)} readOnly /></label>
                          <label className={remainingAmount > 0 ? 'manager-order-profit-negative' : 'manager-order-profit-positive'}>Довг<input value={money(remainingAmount)} readOnly /></label>
                          <label>Статус<input value={paymentVisualStatus} readOnly /></label>
                          <label>Тип оплати<input value={lastPaymentMethod} readOnly /></label>
                          <label>Дата оплати<input value={lastPaymentDate} readOnly /></label>
                          <label>Статус платежу<input value={lastPaymentStatus} readOnly /></label>
                          <label className={orderProfit >= 0 ? 'manager-order-profit-positive' : 'manager-order-profit-negative'}>Прибуток<input value={money(orderProfit)} readOnly /></label>
                          {!linkedContract && canTakePayment && (
                            <>
                              <label>Сума платежу<input type="number" min={0} value={paymentDraft.amount} onChange={(event) => setManagerPaymentDrafts((current) => ({ ...current, [order.id]: { ...managerPaymentDraft(order), amount: event.target.value } }))} /></label>
                              <label>Тип платежу<select value={paymentDraft.type} onChange={(event) => setManagerPaymentDrafts((current) => ({ ...current, [order.id]: { ...managerPaymentDraft(order), type: event.target.value as SimpleLedgerPaymentKind } }))}><option value="предоплата">предоплата</option><option value="оплата">оплата</option><option value="частичная">частичная</option></select></label>
                            </>
                          )}
                        </div>
                        {pendingPayments.length > 0 && (
                          <div className="manager-order-parts-list">
                            {pendingPayments.map((payment) => (
                              <div key={payment.id} className="manager-order-part-row">
                                <span>{money(payment.amount)}</span>
                                <span>{payment.type}</span>
                                <span>{payment.method}</span>
                                <button type="button" onClick={() => props.confirmSimpleManagerPayment(payment.id)}>Підтвердити</button>
                              </div>
                            ))}
                          </div>
                        )}
                      </div>

                      <div className="manager-order-payment manager-order-section">
                        <div className="panel-heading">
                          <h2>Клієнт</h2>
                          <button type="button" onClick={() => props.sendClientNotification(order, notifyEvent, 'SMS', true)}>Відправити повідомлення</button>
                        </div>
                        <div className="manager-order-bottom-meta">
                          <div className="manager-order-field">
                            <strong>Телефон</strong>
                            <div><a href={`tel:${order.phone}`}>{order.phone}</a></div>
                          </div>
                          <div className="manager-order-field">
                            <strong>Останнє повідомлення</strong>
                            <div>{order.notificationHistory?.[0] ? `${notificationDisplay(order.notificationHistory[0].event)} · ${order.notificationHistory[0].status}` : 'Ще не надсилали'}</div>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>

                  <div className="manager-order-details manager-order-section">
                    <details className="manager-order-collapse">
                      <summary>Деталі та собівартість</summary>
                      {addingPartOrderId === order.id ? (
                        <div className="manager-order-part-picker">
                          <input value={partSearch} onKeyDown={handleManagerFieldKeyDown} onChange={(event) => setPartSearch(event.target.value)} placeholder="Пошук деталі" />
                          <div className="manager-order-part-results">
                            {filteredProducts.map((product) => (
                              <button key={product.id} type="button" onClick={() => setSelectedPartProductId(product.id)} className={selectedPartProductId === product.id ? 'manager-order-part-option is-selected' : 'manager-order-part-option'}>
                                {product.name} · {available(product)} · {money(product.price || product.cost)}
                              </button>
                            ))}
                          </div>
                          <div className="manager-order-part-actions">
                            <input type="number" min={1} value={selectedPartQty} onKeyDown={handleManagerFieldKeyDown} onChange={(event) => setSelectedPartQty(event.target.value)} placeholder="Кількість" />
                            <button type="button" className="submit-button" onClick={() => submitManagerPart(order.id)} disabled={!selectedPartProductId}>Додати</button>
                            <button type="button" onClick={() => setAddingPartOrderId('')}>Скасувати</button>
                          </div>
                        </div>
                      ) : (
                        <div className="manager-order-inline-actions">
                          <button type="button" onClick={() => { setAddingPartOrderId(order.id); setPartSearch(''); setSelectedPartProductId(''); setSelectedPartQty('1'); }}>Додати деталь</button>
                        </div>
                      )}
                      <div className="manager-order-parts-list">
                        {order.parts.map((part) => {
                          const partProduct = props.products.find((product) => product.id === part.productId);
                          return (
                            <div key={part.id} className="manager-order-part-row">
                              <span>{partProduct?.name ?? part.productId}</span>
                              <span>{part.qty}</span>
                              <span>{part.batchAllocations?.map((item) => item.batchId).join(', ') || '—'}</span>
                              <span>{money(part.cost)}</span>
                              <button type="button" onClick={() => setManagerExceptionDraft({ orderId: order.id, mode: 'part-return', reason: '', comment: '', amount: '', partId: part.id })}>Повернення</button>
                            </div>
                          );
                        })}
                      </div>
                      {order.parts.length === 0 && <div className="empty-state manager-order-empty-parts">Деталей поки немає</div>}
                      <div className="client-meta manager-order-cost-total">
                        <span>Собівартість</span>
                        <span>{money(orderCost)}</span>
                        <span>Договір: {linkedContract ? `${linkedContract.id}${linkedAct ? ` · ${linkedAct.id}` : ''}` : 'Без договору'}</span>
                      </div>
                    </details>

                    <details className="manager-order-collapse">
                      <summary>Історія та аудит</summary>
                      <div className="manager-order-history-list">
                        {(order.activityLog?.slice(0, 8) ?? []).map((entry) => (
                          <div key={entry.id} className="manager-order-history-row">
                            <span>{entry.action}</span>
                            <small>{entry.detail ? `${entry.detail} · ${entry.date}` : entry.date}</small>
                          </div>
                        ))}
                        {(!order.activityLog || order.activityLog.length === 0) && <div className="manager-order-history-row"><span>Подій ще немає</span></div>}
                      </div>
                    </details>

                    <div className="manager-order-inline-actions">
                      {canCancelOrder && <button type="button" onClick={() => setManagerExceptionDraft({ orderId: order.id, mode: 'cancel', reason: '', comment: '', amount: '' })}>Скасувати</button>}
                      {canReturnToWork && <button type="button" onClick={() => setManagerExceptionDraft({ orderId: order.id, mode: 'reopen', reason: '', comment: '', amount: '' })}>Повернути в роботу</button>}
                      {canRefundMoney && <button type="button" onClick={() => setManagerExceptionDraft({ orderId: order.id, mode: 'refund', reason: '', comment: '', amount: String(paidAmount) })}>Повернення коштів</button>}
                    </div>

                    {managerExceptionDraft?.orderId === order.id && (
                      <div className="manager-order-exception-box">
                        <strong>
                          {managerExceptionDraft.mode === 'cancel'
                            ? 'Скасування замовлення'
                            : managerExceptionDraft.mode === 'reopen'
                              ? 'Повернення в роботу'
                              : managerExceptionDraft.mode === 'refund'
                                ? 'Повернення коштів'
                                : 'Повернення деталі'}
                        </strong>
                        <div className="manager-order-exception-grid">
                          {managerExceptionDraft.mode === 'refund' && (
                            <label className="manager-order-field">
                              <span>Сума</span>
                              <input type="number" min={0} value={managerExceptionDraft.amount} onChange={(event) => setManagerExceptionDraft((current) => current ? { ...current, amount: event.target.value } : current)} />
                            </label>
                          )}
                          <label className="manager-order-field">
                            <span>Причина</span>
                            <input value={managerExceptionDraft.reason} onChange={(event) => setManagerExceptionDraft((current) => current ? { ...current, reason: event.target.value } : current)} placeholder="Обов'язково" />
                          </label>
                          <label className="manager-order-field">
                            <span>Коментар</span>
                            <input value={managerExceptionDraft.comment} onChange={(event) => setManagerExceptionDraft((current) => current ? { ...current, comment: event.target.value } : current)} placeholder="Короткий коментар" />
                          </label>
                        </div>
                        <div className="manager-order-exception-actions">
                          <button type="button" className="submit-button" onClick={() => submitManagerException(order)}>
                            {managerExceptionDraft.mode === 'cancel' ? 'Підтвердити скасування' : managerExceptionDraft.mode === 'reopen' ? 'Повернути в роботу' : managerExceptionDraft.mode === 'refund' ? 'Оформити повернення' : 'Повернути на склад'}
                          </button>
                          <button type="button" onClick={resetManagerException}>Скасувати</button>
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              );
            })()}
            </section>
          </div>
        )}
        </>
        )}

        {managerPaymentModal && (
          <div className="manager-payment-modal-backdrop" onClick={() => setManagerPaymentModal(null)}>
            <section className="manager-payment-modal" onClick={(event) => event.stopPropagation()}>
              <div className="panel-heading">
                <h2>Оплата замовлення {managerPaymentModal.orderId}</h2>
                <span>Каса / термінал / безготівка</span>
              </div>
              <div className="manager-payment-modal-body">
                <label className="manager-payment-field">
                  <span>Сума до оплати</span>
                  <input value={money(managerPaymentModal.due)} readOnly />
                </label>
                <label className="manager-payment-field">
                  <span>Сума оплати</span>
                  <input
                    ref={managerPaymentAmountInputRef}
                    type="number"
                    min={0}
                    max={managerPaymentModal.due}
                    value={managerPaymentModal.amount}
                    onChange={(event) => setManagerPaymentModal((current) => current ? { ...current, amount: event.target.value } : current)}
                  />
                </label>
                <label className="manager-payment-field">
                  <span>Спосіб оплати</span>
                  <select
                    value={managerPaymentModal.method}
                    onChange={(event) => setManagerPaymentModal((current) => current ? { ...current, method: event.target.value as 'наличные' | 'карта' | 'перевод' } : current)}
                  >
                    <option value="наличные">Готівка</option>
                    <option value="карта">Термінал</option>
                    <option value="перевод">Безготівка (IBAN)</option>
                  </select>
                </label>
              </div>
              <div className="action-row manager-payment-actions">
                <button type="button" className="submit-button" onClick={confirmManagerPaymentModal}>Підтвердити</button>
                <button type="button" onClick={() => setManagerPaymentModal(null)}>Скасувати</button>
              </div>
            </section>
          </div>
        )}
        {managerIssueModal && (
          <div className="manager-payment-modal-backdrop" onClick={() => setManagerIssueModal(null)}>
            <section className="manager-payment-modal manager-issue-modal" onClick={(event) => event.stopPropagation()}>
              <div className="panel-heading">
                <h2>Підтвердити видачу замовлення?</h2>
                <span>Підтвердження дії</span>
              </div>
              <div className="manager-payment-modal-body">
                <p className="manager-issue-modal-text">Видати замовлення {managerIssueModal.orderId} клієнту?</p>
              </div>
              <div className="action-row manager-payment-actions">
                <button type="button" className="submit-button manager-issue-confirm-button" onClick={confirmManagerIssueModal}>Видати</button>
                <button type="button" onClick={() => setManagerIssueModal(null)}>Скасувати</button>
              </div>
            </section>
          </div>
        )}
      </div>
    );
  }

  const orderPriorityWeight = (order: ServiceOrder) => {
    if (order.urgent) return 4;
    const priority = orderPriority(order);
    if (priority === 'Високий') return 3;
    if (priority === 'Середній') return 2;
    return 1;
  };
  const byPriorityAndAge = (a: ServiceOrder, b: ServiceOrder) => {
    const priorityDiff = orderPriorityWeight(b) - orderPriorityWeight(a);
    if (priorityDiff !== 0) return priorityDiff;
    return daysSince(b.intakeDate) - daysSince(a.intakeDate);
  };
  const engineerNew = roleOrders.filter((order) => !order.engineerAcceptedAt && !['Готовий до видачі', 'Не підлягає ремонту', 'Видано', 'Закрито', 'Скасовано'].includes(order.status));
  const engineerInWork = roleOrders.filter((order) => ['На діагностиці', 'В ремонті'].includes(order.status));
  const engineerWaitingParts = roleOrders.filter((order) => order.status === 'Очікує запчастину');
  const engineerOnTest = roleOrders.filter((order) => order.status === 'На тестуванні');
  const engineerFinished = roleOrders.filter((order) => ['Готовий до видачі', 'Не підлягає ремонту', 'Видано', 'Закрито'].includes(order.status));
  const engineerTakenOrders = [...roleOrders.filter((order) => order.locationStatus === 'У інженера')].sort((a, b) => {
    const takenDiff = hoursSince(b.takenInRepairAt ?? b.statusChangedAt) - hoursSince(a.takenInRepairAt ?? a.statusChangedAt);
    if (takenDiff !== 0) return takenDiff;
    return byPriorityAndAge(a, b);
  });
  const managerReadyOrders = [...props.orders.filter((order) => ['Готовий до видачі', 'Не підлягає ремонту', 'Очікує оплати', 'Очікує клієнта'].includes(order.status))].sort((a, b) => {
    const returnedDiff = (b.returnedToCellAt ? 1 : 0) - (a.returnedToCellAt ? 1 : 0);
    if (returnedDiff !== 0) return returnedDiff;
    if (a.returnedToCellAt && b.returnedToCellAt) {
      return daysSince(a.returnedToCellAt) - daysSince(b.returnedToCellAt);
    }
    return byPriorityAndAge(a, b);
  });
  const managerInboxOrders = props.orders.filter((order) => ['Прийнято', 'Очікує погодження', 'Погоджено', 'Готовий до видачі', 'Очікує оплати', 'Очікує клієнта'].includes(order.status));
  const engineerWorkbenchOrders = [...engineerNew, ...engineerInWork, ...engineerWaitingParts, ...engineerOnTest, ...engineerFinished]
    .filter((order, index, list) => list.findIndex((item) => item.id === order.id) === index)
    .sort(byPriorityAndAge);
  const supervisorWatchOrders = [...props.orders].sort((a, b) => {
    const aClosed = ['Видано', 'Закрито', 'Скасовано'].includes(a.status) ? 1 : 0;
    const bClosed = ['Видано', 'Закрито', 'Скасовано'].includes(b.status) ? 1 : 0;
    if (aClosed !== bClosed) return aClosed - bClosed;
    return daysSince(b.statusChangedAt) - daysSince(a.statusChangedAt);
  });
  const managerStatusOptions: OrderStatus[] = ['Готовий до видачі', 'Очікує оплати', 'Очікує клієнта', 'Не підлягає ремонту'];
  const primaryOrders = isEngineer ? engineerNew : isManager ? managerInboxOrders : isSupervisor ? supervisorWatchOrders : props.orders;
  const primaryListTitle = isEngineer ? 'Доступні замовлення' : isManager ? 'Замовлення менеджера' : isSupervisor ? 'Усі замовлення під контролем' : 'Журнал ремонтів';
  const primaryListHint = isEngineer ? 'взяти в ремонт без ручного вибору статусів' : isManager ? 'прийом, клієнт, статус, видача' : isSupervisor ? 'перегляд і контроль без зайвих дій' : 'актуальні замовлення';
  const supervisorActiveOrders = props.orders.filter((order) => !['Видано', 'Закрито', 'Скасовано'].includes(order.status));
  const supervisorReadyOrders = props.orders.filter((order) => ['Готовий до видачі', 'Очікує оплати', 'Очікує клієнта', 'Не підлягає ремонту'].includes(order.status));
  const supervisorDebtOrders = props.orders.filter((order) => orderTotals(order).debt > 0);
  return (
    <div className="page-grid">
      <PageTitle eyebrow={isEngineer ? 'Мої замовлення' : 'CRM сервісного центру'} title={isEngineer ? `Робоче місце інженера: ${props.activeUser.name}` : 'Замовлення сервісного центру'} text={`Роль: ${roleDisplay(props.activeUser.role)}. На екрані лише основні дії по ремонту, статусах, клієнту та оплаті без втрати поточного функціоналу.`} />
      {isEngineer && (
        <section className="panel">
          <div className="panel-heading"><h2>Замовлення у інженера</h2><span>{engineerTakenOrders.length} у роботі зараз</span></div>
          <div className="order-list">
            {engineerTakenOrders.map((order) => (
              <article
                className="order-card"
                key={`taken-${order.id}`}
                onClick={() => props.setSelectedOrderId(order.id)}
                style={hoursSince(order.takenInRepairAt ?? order.statusChangedAt) > ENGINEER_OVERDUE_HOURS ? { borderColor: '#dc2626', boxShadow: '0 0 0 2px rgba(220, 38, 38, 0.15)', cursor: 'pointer' } : order.urgent ? { borderColor: '#f59e0b', boxShadow: '0 0 0 2px rgba(245, 158, 11, 0.15)', cursor: 'pointer' } : { cursor: 'pointer' }}
              >
                <div className="order-title"><strong>{order.id}</strong></div>
                <h3>{order.device}</h3>
                <div className="order-meta-row">
                  <span>Комірка: {order.locationCode ?? 'не призначено'}</span>
                  <span>Взято: {order.takenInRepairAt ?? 'не зафіксовано'}</span>
                  <span>Таймер: {engineerElapsedLabel(order.takenInRepairAt ?? order.statusChangedAt)}</span>
                </div>
                <div className="order-meta-row">
                  {order.urgent && <span style={{ color: '#b45309', fontWeight: 700 }}>ТЕРМІНОВО</span>}
                  {hoursSince(order.takenInRepairAt ?? order.statusChangedAt) > ENGINEER_OVERDUE_HOURS && <span style={{ color: '#b91c1c', fontWeight: 700 }}>ПРОСТРОЧЕНО</span>}
                </div>
                <div className="action-row">
                  <button type="button" onClick={(event) => { event.stopPropagation(); props.returnOrderToCellReady(order); }} disabled={!order.locationCode || order.locationStatus !== 'У інженера'}>Готово, повернути в комірку</button>
                </div>
              </article>
            ))}
            {engineerTakenOrders.length === 0 && <div className="empty-state">Зараз у інженера немає замовлень на руках.</div>}
          </div>
        </section>
      )}
      {isManager && (
        <QuickOrderIntake
          phone={props.quickPhone}
          clientName={props.quickClientName}
          device={props.quickDevice}
          serial={props.quickSerial}
          problem={props.quickProblem}
          appearance={props.quickAppearance}
          estimatedAmount={props.quickEstimatedAmount}
          engineerId={props.quickEngineerId}
          contractId={props.quickContractId}
          locationCode={props.quickLocationCode}
          comment={props.quickComment}
          managerName={props.activeUser.name}
          users={props.users}
          contracts={props.contracts}
          customerList={props.customerList}
          warehouseLocations={props.warehouseLocations}
          onPhoneChange={props.setQuickPhone}
          onClientNameChange={props.setQuickClientName}
          onDeviceChange={props.setQuickDevice}
          onSerialChange={props.setQuickSerial}
          onProblemChange={props.setQuickProblem}
          onAppearanceChange={props.setQuickAppearance}
          onEstimatedAmountChange={props.setQuickEstimatedAmount}
          onEngineerChange={props.setQuickEngineerId}
          onContractChange={props.setQuickContractId}
          onLocationChange={props.setQuickLocationCode}
          onCommentChange={props.setQuickComment}
          onSubmit={props.createQuickOrder}
        />
      )}
      {isManager && (
        <section className="stats-grid">
          <Metric icon={<ClipboardList />} label="Прийом" value={String(managerInboxOrders.filter((order) => ['Прийнято', 'Очікує погодження', 'Погоджено'].includes(order.status)).length)} hint="нові та на погодженні" />
          <Metric icon={<Users />} label="Клієнт" value={String(managerInboxOrders.filter((order) => ['Очікує клієнта', 'Очікує оплати'].includes(order.status)).length)} hint="очікують зв'язок та оплату" />
          <Metric icon={<PackageCheck />} label="До видачі" value={String(managerReadyOrders.length)} hint="готові для менеджера" />
          <Metric icon={<Banknote />} label="З боргом" value={String(managerInboxOrders.filter((order) => orderTotals(order).debt > 0).length)} hint="потрібна оплата" />
        </section>
      )}
      {isSupervisor && (
        <section className="stats-grid">
          <Metric icon={<ClipboardList />} label="Все заказы" value={String(props.orders.length)} hint="общий контроль сервиса" />
          <Metric icon={<Wrench />} label="Активные" value={String(supervisorActiveOrders.length)} hint="ещё не завершены" />
          <Metric icon={<PackageCheck />} label="Готово" value={String(supervisorReadyOrders.length)} hint="ожидают выдачу" />
          <Metric icon={<Banknote />} label="С оплатой" value={String(supervisorDebtOrders.length)} hint="есть долг клиента" />
        </section>
      )}
      {isManager && (
        <section className="panel">
          <div className="panel-heading"><h2>Готово до видачі</h2><span>{managerReadyOrders.length} потребують дій менеджера</span></div>
          <div className="order-list">
         {managerReadyOrders.map((order) => {
           const totals = orderTotals(order);
           const orderControl = buildOrderControlState(order, props.allOrders, props.allNotifications.filter((item) => item.orderId === order.id), props.orderUnits);
           const issueBlocked = orderControl.critical.length > 0 || totals.debt > 0 || !order.locationCode || order.locationStatus === 'У інженера';
           const availableStatuses = managerStatusOptions.filter((status) => status === order.status || (orderStatusFlow[order.status] ?? []).includes(status));
           return (
             <div key={order.id} className={`order-card ${props.selectedOrderId === order.id ? 'order-card-active' : ''}`} style={order.returnedToCellAt ? { borderColor: '#16a34a', boxShadow: '0 0 0 2px rgba(22, 163, 74, 0.18)' } : undefined}>
               <div className="order-title"><strong>{order.id}</strong><span className={statusClass[order.status]}>{statusDisplay(order.status)}</span></div>
               <h3>{order.device}</h3>
               <p>{order.client} · {order.issue}</p>
               <p>Менеджер: {getOrderCreatorName(order, props.users)}</p>
               {order.returnedToCellAt && <p style={{ color: '#166534', fontWeight: 700 }}>Готово до видачі: в комірці {order.locationCode}</p>}
               <div className="order-meta-row">
                 <span>{simpleRepairPaymentStatus(order)}</span>
                 <span>{money(totals.total)}</span>
                  <span>{order.locationCode ?? 'без комірки'}</span>
                  <span>борг {money(totals.debt)}</span>
               </div>
               <div className="order-meta-row">
                 {order.urgent && <span style={{ color: '#b45309', fontWeight: 700 }}>ТЕРМІНОВО</span>}
                 {order.locationStatus === 'У інженера' && <span style={{ color: '#b91c1c', fontWeight: 700 }}>ЩЕ У ІНЖЕНЕРА</span>}
               </div>
               <div className="order-meta-row">
                 <button type="button" onClick={() => props.setSelectedOrderId(order.id)}>Відкрити замовлення</button>
                 <select
                    aria-label={`Статус ${order.id}`}
                   value={order.status}
                   onChange={(e) => {
                     const newStatus = e.target.value as OrderStatus;
                     if (newStatus === order.status) return;
                     const commentByStatus: Partial<Record<OrderStatus, string>> = {
                        'Готовий до видачі': 'Менеджер повернув замовлення у статус готовності до видачі.',
                        'Очікує оплати': 'Менеджер очікує оплату перед видачею.',
                       'Очікує клієнта': 'Менеджер зв’язався з клієнтом і очікує видачу.',
                        'Не підлягає ремонту': 'Менеджер перевів замовлення у статус неможливого ремонту.',
                     };
                      props.changeOrderStatus(order, newStatus, commentByStatus[newStatus] ?? 'Менеджер змінив статус у блоці готових до видачі.');
                   }}
                 >
                   {availableStatuses.map((status) => (
                      <option key={status} value={status}>{statusDisplay(status)}</option>
                   ))}
                 </select>
                  <button type="button" onClick={() => {
                    if (orderControl.warnings.length > 0) props.logRiskConfirmation(order, 'Менеджер видав замовлення зі списку готових', orderControl.warnings.map((item) => item.message));
                    props.issueReadyOrder(order);
                  }} disabled={issueBlocked}>Видати</button>
               </div>
             </div>
           );
         })}
            {managerReadyOrders.length === 0 && <div className="empty-state">Зараз немає замовлень, готових до видачі.</div>}
          </div>
        </section>
      )}
      {isManager ? (
        <>
          <section className="panel manager-orders-stage">
            <div className="panel-heading"><h2>{primaryListTitle}</h2><span>{primaryOrders.length} · {primaryListHint}</span></div>
            <OrderSummaryList orders={primaryOrders} selectedId={props.selectedOrderId} onSelect={props.setSelectedOrderId} showPriority users={props.users} />
          </section>
          <OrderDetail {...props} />
        </>
      ) : (
        <section className="content-split orders-layout">
          <div className="panel">
            <div className="panel-heading"><h2>{primaryListTitle}</h2><span>{primaryOrders.length} · {primaryListHint}</span></div>
            <OrderSummaryList orders={primaryOrders} selectedId={props.selectedOrderId} onSelect={props.setSelectedOrderId} showPriority={isEngineer || isSupervisor} users={props.users} />
          </div>
          <OrderDetail {...props} />
        </section>
      )}
    </div>
  );
}

function MyOrdersPage({ orders, activeUser, acceptOrderWork, changeOrderStatus, returnOrderToCellReady }: { orders: ServiceOrder[]; activeUser: User; acceptOrderWork: (order: ServiceOrder) => void; changeOrderStatus: (order: ServiceOrder, nextStatus: OrderStatus, comment: string) => void; returnOrderToCellReady: (order: ServiceOrder) => void }) {
  const engineerOrders = [...orders].sort((a, b) => {
    const aScore = (!a.engineerAcceptedAt ? 300 : 0)
      + (['На діагностиці', 'В ремонті'].includes(a.status) ? 220 : 0)
      + (a.status === 'Очікує запчастину' ? 160 : 0)
      + (a.status === 'На тестуванні' ? 120 : 0)
      + Math.min(daysSince(a.statusChangedAt), 30);
    const bScore = (!b.engineerAcceptedAt ? 300 : 0)
      + (['На діагностиці', 'В ремонті'].includes(b.status) ? 220 : 0)
      + (b.status === 'Очікує запчастину' ? 160 : 0)
      + (b.status === 'На тестуванні' ? 120 : 0)
      + Math.min(daysSince(b.statusChangedAt), 30);
    return bScore - aScore || (parseDateTime(b.statusChangedAt)?.getTime() ?? 0) - (parseDateTime(a.statusChangedAt)?.getTime() ?? 0);
  });

  return (
    <div className="page-grid">
      <PageTitle eyebrow="Мої замовлення" title={`Інженер: ${activeUser.name}`} text="Тільки мої замовлення і робочі дії." />
      <section className="panel">
        <div className="panel-heading"><h2>Мої замовлення</h2><span>{engineerOrders.length}</span></div>
        <div className="order-list">
          {engineerOrders.map((order) => {
            const canStart = !order.engineerAcceptedAt;
            const canTest = ['На діагностиці', 'В ремонті'].includes(order.status);
            const canFinish = ['На діагностиці', 'В ремонті', 'На тестуванні'].includes(order.status);
            const isHot = !order.engineerAcceptedAt || hoursSince(order.statusChangedAt) >= ENGINEER_OVERDUE_HOURS;
            return (
              <article className={`order-card${isHot ? ' order-card-active' : ''}`} key={order.id}>
                <div className="order-title">
                  <strong>{order.id}</strong>
                  <span className={statusClass[order.status]}>{statusDisplay(order.status)}</span>
                </div>
                <h3>{order.device}</h3>
                <p>{order.issue}</p>
                <div className="order-meta-row">
                  <span>{order.client}</span>
                  <span>{order.locationCode ?? 'без комірки'}</span>
                  <span>{engineerElapsedLabel(order.takenInRepairAt ?? order.statusChangedAt)}</span>
                </div>
                <div className="action-row">
                  {canStart && <button type="button" onClick={() => acceptOrderWork(order)}>Почати</button>}
                  {canTest && <button type="button" onClick={() => changeOrderStatus(order, 'На тестуванні', 'Інженер передав пристрій на тестування.')}>На тест</button>}
                  {canFinish && <button type="button" onClick={() => returnOrderToCellReady(order)}>Готово</button>}
                </div>
              </article>
            );
          })}
          {engineerOrders.length === 0 && <div className="empty-state">Замовлень для роботи зараз немає.</div>}
        </div>
      </section>
    </div>
  );
}

function OrderDetail({ selectedOrder, allOrders, orderUnits, warehouseLocations, orderMovementLogs, moveOrder, products, selectedProductId, setSelectedProductId, qty, setQty, addPartToRepair, orderPart, reserveArrived, issueToEngineer, markInstalled, returnServicePart, addOrderPayment, changeOrderStatus, acceptOrderWork, returnOrderToCellReady, reassignEngineer, closeOrder, issueReadyOrder, oneClickManagerIssue, transferOrderToBas, printDocument, createServiceOrderDocument, printServiceOrderDocument, documents, taxInvoices, notifications, allNotifications, sendClientNotification, createNotificationDraft, approval, sendRepairApproval, recordApprovalResponse, markApprovalNoAnswer, logRiskConfirmation, ensureOrderDocumentRecord, logOrderDocumentPrint, signOrderAct, createTaxInvoiceForOrder, registerTaxInvoice, markEngineerWorkCompleted, cancelOrder, refundOrder, cancelOrderAct, reopenOrder, orderVersions, suggestedDocumentAction, clearSuggestedDocumentAction, dashboardFocus, clearDashboardFocus, notifyUser, canDo, showCost, activeUser, users }: {
  selectedOrder: ServiceOrder;
  allOrders: ServiceOrder[];
  orderUnits: OrderUnit[];
  warehouseLocations: WarehouseLocation[];
  orderMovementLogs: OrderMovementLog[];
  moveOrder: (orderId: string, newLocation: string) => void;
  products: Product[];
  selectedProductId: string;
  setSelectedProductId: (id: string) => void;
  qty: number;
  setQty: (qty: number) => void;
  addPartToRepair: () => void;
  orderPart: (order: ServiceOrder, part: OrderPart) => void;
  reserveArrived: (order: ServiceOrder, part: OrderPart) => void;
  issueToEngineer: (order: ServiceOrder, part: OrderPart) => void;
  markInstalled: (order: ServiceOrder, part: OrderPart) => void;
  returnServicePart: (order: ServiceOrder, part: OrderPart, destination: 'На склад' | 'Брак') => void;
  addOrderPayment: (order: ServiceOrder, amount: number, method: PaymentMethod, comment: string) => void;
  changeOrderStatus: (order: ServiceOrder, nextStatus: OrderStatus, comment: string) => void;
  acceptOrderWork: (order: ServiceOrder) => void;
  returnOrderToCellReady: (order: ServiceOrder) => void;
  reassignEngineer: (order: ServiceOrder) => void;
  closeOrder: (order: ServiceOrder) => void;
  issueReadyOrder: (order: ServiceOrder) => void;
  oneClickManagerIssue: (order: ServiceOrder, payment?: { method: PaymentMethod; amount?: number; partial?: boolean }) => void;
  transferOrderToBas: (order: ServiceOrder) => void;
  printDocument: (kind: DocumentKind, entityType: PrintDocument['entityType'], entityId: string, clientOrSupplier: string) => void;
  createServiceOrderDocument: (kind: ServiceOrderDocumentKind, order: ServiceOrder) => PrintDocument | undefined;
  printServiceOrderDocument: (kind: ServiceOrderDocumentKind, order: ServiceOrder) => void;
  documents: PrintDocument[];
  taxInvoices: TaxInvoice[];
  notifications: ClientNotification[];
  allNotifications: ClientNotification[];
  sendClientNotification: (order: ServiceOrder, event: NotificationEvent, preferredChannel?: NotificationChannel, manual?: boolean, messageOverride?: string) => void;
  createNotificationDraft: (order: ServiceOrder, event: NotificationEvent, channel: NotificationChannel, text: string, status: NotificationStatus) => ClientNotification;
  approval?: RepairApproval;
  sendRepairApproval: (order: ServiceOrder, channel?: NotificationChannel, extra?: { description?: string; amount?: number; comment?: string }) => void;
  recordApprovalResponse: (approvalId: string, accepted: boolean, manual?: boolean) => void;
  markApprovalNoAnswer: (approvalId: string) => void;
  logRiskConfirmation: (order: ServiceOrder, action: string, risks: string[]) => void;
  ensureOrderDocumentRecord: (kind: 'Рахунок на оплату' | 'Акт надання послуг' | 'Наряд на ремонт' | 'Гарантійний талон', order: ServiceOrder) => void;
  logOrderDocumentPrint: (order: ServiceOrder, kind: 'Рахунок на оплату' | 'Акт надання послуг' | 'Наряд на ремонт' | 'Гарантійний талон') => void;
  signOrderAct: (order: ServiceOrder) => void;
  createTaxInvoiceForOrder: (orderId: string) => void;
  registerTaxInvoice: (orderId: string) => void;
  markEngineerWorkCompleted: (order: ServiceOrder) => void;
  cancelOrder: (order: ServiceOrder) => void;
  refundOrder: (order: ServiceOrder) => void;
  cancelOrderAct: (order: ServiceOrder) => void;
  reopenOrder: (order: ServiceOrder) => void;
  orderVersions: OrderVersion[];
  suggestedDocumentAction: SuggestedDocumentAction | null;
  clearSuggestedDocumentAction: () => void;
  dashboardFocus: { orderId: string; target: DashboardFocusTarget } | null;
  clearDashboardFocus: () => void;
  notifyUser: (message: string) => void;
  canDo: (permission: Permission) => boolean;
  showCost: boolean;
  activeUser: User;
  users: User[];
}) {
  const totals = orderTotals(selectedOrder);
  const simplePaymentStatus = simpleRepairPaymentStatus(selectedOrder);
  const controlState = buildOrderControlState(selectedOrder, allOrders, notifications, orderUnits);
  const statusAge = daysSince(selectedOrder.statusChangedAt);
  const totalAge = daysSince(selectedOrder.intakeDate);
  const nextStatuses = orderStatusFlow[selectedOrder.status] ?? [];
  const isEngineerRole = activeUser.role === 'Інженер';
  const isManagerRole = activeUser.role === 'Менеджер';
  const isSupervisorRole = activeUser.role === 'Руководитель';
  const isAdminRole = activeUser.role === 'Адміністратор';
  const canSeeFinance = !isEngineerRole;
  const canManageDocuments = isManagerRole || isAdminRole;
  const canSendApproval = isManagerRole || isAdminRole;
  const canAddParts = isEngineerRole || isAdminRole;
  const canManageRepairFlow = isEngineerRole || isManagerRole || isAdminRole;
  const canManageClientWork = isManagerRole || isAdminRole;
  const canShowReadonlyOverview = isSupervisorRole;
  const canShowEngineerView = isEngineerRole;
  const canShowManagerView = isManagerRole;
  const canShowPartActions = isEngineerRole || isAdminRole;
  const canShowRoleFocusBlock = !isEngineerRole && !isSupervisorRole;
  const canShowCommentBlock = !isSupervisorRole;
  const canShowHistoryBlock = !isSupervisorRole;
  const canShowQrPanel = !canShowEngineerView && !canShowManagerView && !canShowReadonlyOverview;
  const canShowDetailedFinance = canSeeFinance && !canShowManagerView && !canShowReadonlyOverview;
  const canShowPartsSection = !canShowReadonlyOverview;
  const [paymentAmount, setPaymentAmount] = useState(0);
  const [paymentMethod, setPaymentMethod] = useState<PaymentMethod>('Готівка');
  const [selectedNotificationEvent, setSelectedNotificationEvent] = useState<NotificationEvent>('Прийом пристрою');
  const [notificationPreview, setNotificationPreview] = useState<{ event: NotificationEvent; text: string; channel: NotificationChannel } | null>(null);
  const [extraWorkDescription, setExtraWorkDescription] = useState('');
  const [extraWorkAmount, setExtraWorkAmount] = useState(0);
  const [extraWorkComment, setExtraWorkComment] = useState('');
  const [isDocumentPreviewOpen, setIsDocumentPreviewOpen] = useState(false);
  const [isLabelPreviewOpen, setIsLabelPreviewOpen] = useState(false);
  const [pendingLocationCode, setPendingLocationCode] = useState(selectedOrder.locationCode ?? '');
  const [selectedDocumentKind, setSelectedDocumentKind] = useState<'Рахунок на оплату' | 'Акт надання послуг' | 'Видаткова накладна' | 'Наряд на ремонт' | 'Гарантійний талон'>('Наряд на ремонт');
  const [documentPreviewNonce, setDocumentPreviewNonce] = useState(0);
  const [isMiniCashOpen, setIsMiniCashOpen] = useState(false);
  const documentIframeRef = React.useRef<HTMLIFrameElement | null>(null);
  const effectivePaymentAmount = Math.min(paymentAmount > 0 ? paymentAmount : totals.debt, totals.debt);
  const recordOrderPayment = (comment = 'Оплата ремонту') => {
    if (totals.works <= 0) {
      notifyUser('Прийняти оплату не можна: не вказана сума робіт.');
      return;
    }
    addOrderPayment(selectedOrder, effectivePaymentAmount, paymentMethod, comment);
    if (effectivePaymentAmount > 0) {
      printDocument('Касовий чек', 'service_order', selectedOrder.id, selectedOrder.client);
    }
  };
  const directOrderUrl = typeof window === 'undefined' ? selectedOrder.qrUrl : `${window.location.origin}${window.location.pathname}#${selectedOrder.qrUrl}`;
  const hasDocument = (kind: DocumentKind) => documents.some((document) => document.entityType === 'service_order' && document.entityId === selectedOrder.id && document.kind === kind);
  const hasIssueAct = ['Акт надання послуг', 'Акт виконаних робіт', 'Акт видачі', 'Акт технічного стану'].some((kind) => hasDocument(kind as DocumentKind));
  const actDocument = documents.find((document) => document.entityType === 'service_order' && document.entityId === selectedOrder.id && document.kind === 'Акт надання послуг');
  const actDocumentStatus = getOrderActStatus(documents, selectedOrder.id);
  const selectedOrderTaxInvoice = [...taxInvoices]
    .filter((invoice) => invoice.orderId === selectedOrder.id)
    .sort((a, b) => (parseDateTime(b.createdAt)?.getTime() ?? 0) - (parseDateTime(a.createdAt)?.getTime() ?? 0))[0];
  const isVatOrder = Boolean(selectedOrder.legalEntity || selectedOrder.vatStatus);
  const dealClosed = isOrderDealClosed(selectedOrder, documents, taxInvoices);
  const documentsOutdated = orderHasDocumentVersionMismatch(selectedOrder, documents);
  const isReadyForClient = ['Готовий до видачі', 'Очікує оплати', 'Очікує клієнта', 'Видано', 'Закрито'].includes(selectedOrder.status);
  const canFormActAndInvoice = canManageDocuments && isReadyForClient;
  const canFormTechnicalAct = canManageDocuments && selectedOrder.status === 'Не підлягає ремонту';
  const canCreateDeliveryNote = canManageDocuments && selectedOrder.parts.length > 0;
  const releaseBlockReason = getBlockReason(selectedOrder, actDocumentStatus);
  const canIssueOrder = canRelease(selectedOrder, actDocumentStatus) && roleCanSetOrderStatus(activeUser.role, 'Видано') && hasIssueAct && Boolean(selectedOrder.locationCode) && selectedOrder.locationStatus !== 'У інженера';
  const canTransferBas = isReadyForClient && hasIssueAct && hasDocument('Рахунок на оплату') && totals.paid > 0 && (selectedOrder.parts.length === 0 || hasDocument('Видаткова накладна'));
  const clientNotified = notifications.some((item) => item.status === 'Відправлено');
  const selectedNotificationStatus = notifications.find((item) => item.event === selectedNotificationEvent && item.status === 'Відправлено');
  const lastOrderNotification = notifications[0] ?? selectedOrder.notificationHistory?.[0];
  const canPreviewOrderDocuments = isManagerRole || isSupervisorRole || isAdminRole;
  const canPreviewInvoice = canPreviewOrderDocuments && totals.total > 0;
  const canPreviewAct = canPreviewOrderDocuments && selectedOrder.status === 'Готовий до видачі';
  const canPreviewWorkOrder = canPreviewOrderDocuments;
  const canPreviewWarranty = canPreviewOrderDocuments && ['Видано', 'Закрито'].includes(selectedOrder.status);
  const activeDocumentSuggestion = suggestedDocumentAction?.orderId === selectedOrder.id ? suggestedDocumentAction : null;
  const activeDashboardFocus = dashboardFocus?.orderId === selectedOrder.id ? dashboardFocus : null;
  const displayedOrderAmount = totals.total > 0 ? money(totals.total) : selectedOrder.estimatedAmount ? `${money(selectedOrder.estimatedAmount)} · орієнтовно` : 'Не визначено';
  const orderCreatorName = getOrderCreatorName(selectedOrder, users);
  const canCompleteRepair = selectedOrder.works.length > 0 && totals.works > 0 && Boolean(selectedOrder.engineerAcceptedAt);
  const managerCanAcceptIntoWork = canShowManagerView && selectedOrder.status === 'Прийнято';
  const managerCanMarkReadyForIssue = canShowManagerView
    && Boolean(selectedOrder.engineerAcceptedAt)
    && Boolean(selectedOrder.engineerWorkCompletedAt)
    && Boolean(selectedOrder.returnedToCellAt)
    && Boolean(selectedOrder.locationCode)
    && selectedOrder.locationStatus === 'У комірці'
    && selectedOrder.status === 'Готовий до видачі';
  const managerCanTakePayment = canShowManagerView && canDo('finance.payments') && totals.works > 0 && totals.debt > 0;
  const managerCanIssueFromCard = canShowManagerView
    && Boolean(selectedOrder.locationCode)
    && selectedOrder.locationStatus !== 'У інженера'
    && canRelease(selectedOrder, actDocumentStatus);
  const managerNextStep = managerCanAcceptIntoWork
    ? 'Передати інженеру'
    : managerCanMarkReadyForIssue
      ? 'Готово до видачі'
      : managerCanTakePayment
        ? 'Прийняти оплату'
        : canIssueOrder
          ? 'Видати замовлення'
          : 'Очікує наступного етапу';
  const taxInvoiceHint = !isVatOrder
    ? 'ПН для цього замовлення не потрібна.'
    : selectedOrderTaxInvoice
      ? `Поточний статус: ${selectedOrderTaxInvoice.number} · ${selectedOrderTaxInvoice.status}`
      : 'Після першої події створіть ПН з цього замовлення.';
  const taxInvoiceBadgeText = !isVatOrder
    ? 'Не потрібна'
    : selectedOrderTaxInvoice
      ? selectedOrderTaxInvoice.status
      : 'Очікує створення';
  const taxInvoiceBadgeClass = !isVatOrder
    ? 'is-neutral'
    : selectedOrderTaxInvoice
      ? 'is-created'
      : 'is-waiting';
  const documentButtonStyle = (backgroundColor: string) => ({
    display: 'inline-flex',
    alignItems: 'center',
    gap: '8px',
    backgroundColor,
    color: '#fff',
    border: '1px solid rgba(15, 23, 42, 0.16)',
  });
  const orderClientDetails = clientDetails(selectedOrder.client, selectedOrder.phone);
  const managerAcceptIntoWork = () => {
    if (selectedOrder.status !== 'Прийнято') {
      notifyUser('Передати інженеру можна тільки нове замовлення у статусі "Прийнято".');
      return;
    }
    changeOrderStatus(selectedOrder, 'На діагностиці', `Менеджер ${activeUser.name} передав замовлення інженеру ${selectedOrder.engineer}.`);
    openDocumentPreview('Наряд на ремонт');
    notifyUser(`Замовлення ${selectedOrder.id} передано інженеру. Наряд відкрито одразу.`);
  };
  const managerMarkReadyForIssue = () => {
    if (selectedOrder.locationStatus === 'У інженера') {
      notifyUser('Готово до видачі недоступно: замовлення ще у інженера.');
      return;
    }
    if (!selectedOrder.engineerWorkCompletedAt || !selectedOrder.returnedToCellAt) {
      notifyUser('Готово до видачі недоступно: інженер ще не завершив ремонт і не повернув замовлення в комірку.');
      return;
    }
    ensureOrderDocumentRecord('Акт надання послуг', selectedOrder);
    if (actDocumentStatus !== 'Підписано') signOrderAct(selectedOrder);
    openDocumentPreview('Акт надання послуг');
    notifyUser(`Замовлення ${selectedOrder.id} готове до видачі та знаходиться в комірці ${selectedOrder.locationCode}.`);
  };
  const handleManagerIssueClick = () => {
    if (!managerCanIssueFromCard) {
      notifyUser('Видати замовлення зараз не можна: перевірте ячейку, акт і стан ремонту.');
      return;
    }
    if (totals.debt > 0) {
      setIsMiniCashOpen(true);
      return;
    }
    oneClickManagerIssue(selectedOrder);
  };
  const currentLocation = warehouseLocations.find((location) => location.code === selectedOrder.locationCode);
  const repairLocations = warehouseLocations.filter((location) => location.zone === 'REPAIR');
  const locationOptions = warehouseLocations.filter((location) => (
    location.code === selectedOrder.locationCode
      || !allOrders.some((order) => order.id !== selectedOrder.id && order.locationCode === location.code)
  ));
  const locationMovements = orderMovementLogs.filter((entry) => entry.orderId === selectedOrder.id).slice(0, 5);
  const labelStatusColor = selectedOrder.status === 'Прийнято'
    ? '#facc15'
    : ['В ремонті', 'На діагностиці', 'На тестуванні'].includes(selectedOrder.status)
      ? '#2563eb'
      : ['Очікує погодження', 'Погоджено'].includes(selectedOrder.status)
        ? '#f97316'
        : ['Готовий до видачі', 'Очікує клієнта', 'Очікує оплати'].includes(selectedOrder.status)
          ? '#16a34a'
          : ['Не підлягає ремонту', 'Скасовано'].includes(selectedOrder.status)
            ? '#dc2626'
            : '#64748b';
  const labelStatusText = ['Готовий до видачі', 'Очікує клієнта', 'Очікує оплати'].includes(selectedOrder.status)
    ? 'ГОТОВО'
    : ['Очікує погодження', 'Очікує запчастину', 'Пауза / відкладено'].includes(selectedOrder.status)
      ? 'ОЧІКУВАННЯ'
      : 'РЕМОНТ';
  const labelLocationStateText = selectedOrder.locationStatus === 'У інженера' ? 'У ІНЖЕНЕРА' : selectedOrder.locationStatus === 'У комірці' ? 'В ЯЧЕЙЦІ' : 'ПОЗА ЯЧЕЙКОЮ';
  const documentNumberByKind: Record<'Рахунок на оплату' | 'Акт надання послуг' | 'Видаткова накладна' | 'Наряд на ремонт' | 'Гарантійний талон', string> = {
    'Рахунок на оплату': documents.find((document) => document.kind === 'Рахунок на оплату' && document.entityId === selectedOrder.id)?.number ?? `${documentPrefix['Рахунок на оплату']}-${selectedOrder.id.replace(/\D/g, '')}`,
    'Акт надання послуг': documents.find((document) => document.kind === 'Акт надання послуг' && document.entityId === selectedOrder.id)?.number ?? `${documentPrefix['Акт надання послуг']}-${selectedOrder.id.replace(/\D/g, '')}`,
    'Видаткова накладна': documents.find((document) => document.kind === 'Видаткова накладна' && document.entityId === selectedOrder.id)?.number ?? `${documentPrefix['Видаткова накладна']}-${selectedOrder.id.replace(/\D/g, '')}`,
    'Наряд на ремонт': documents.find((document) => document.kind === 'Наряд на ремонт' && document.entityId === selectedOrder.id)?.number ?? `${documentPrefix['Наряд на ремонт']}-${selectedOrder.id.replace(/\D/g, '')}`,
    'Гарантійний талон': documents.find((document) => document.kind === 'Гарантійний талон' && document.entityId === selectedOrder.id)?.number ?? `${documentPrefix['Гарантійний талон']}-${selectedOrder.id.replace(/\D/g, '')}`,
  };
  const issueCriticalMessages = [
    ...controlState.critical.map((item) => item.message),
    ...(releaseBlockReason === 'Є борг' ? ['Нельзя выдать заказ без оплаты.'] : []),
    ...(releaseBlockReason === 'Платіж не підтверджено' ? [pendingPaymentReason(selectedOrder) || 'Платіж не підтверджено.'] : []),
    ...(releaseBlockReason === 'Акт не підписано' ? ['Неможливо видати замовлення без підписаного акта.'] : []),
    ...(!selectedOrder.locationCode || selectedOrder.locationStatus === 'У інженера' ? ['Перед видачею замовлення має бути повернене у закріплену ячейку.'] : []),
  ];
  const acceptWorkCriticalMessages = [
    ...controlState.critical.map((item) => item.message),
    ...(!selectedOrder.engineer ? ['Нельзя взять заказ в работу без назначенного инженера.'] : []),
  ];
  const statusActionCriticalMessages = controlState.critical.map((item) => item.message);
  const warningMessages = controlState.warnings.map((item) => item.message);
  const runWithWarnings = (actionLabel: string, fn: () => void, extraWarnings: string[] = []) => {
    const risks = [...warningMessages, ...extraWarnings];
    if (risks.length > 0) {
      logRiskConfirmation(selectedOrder, actionLabel, risks);
    }
    fn();
  };
  const openNotificationPreview = (event = selectedNotificationEvent) => {
    setSelectedNotificationEvent(event);
    const templateChannel = notificationTemplateForEvent(event)?.channel ?? 'SMS';
    const text = renderNotificationText(selectedOrder, event);
    createNotificationDraft(selectedOrder, event, templateChannel, text, 'Створено');
    setNotificationPreview({ event, text, channel: templateChannel });
  };
  const openExtraApprovalPreview = () => {
    if (!extraWorkDescription.trim() || extraWorkAmount <= 0) {
      notifyUser('Для погодження доп. робіт заповніть опис і суму.');
      return;
    }
    const text = renderNotificationText(selectedOrder, 'Потрібне погодження додаткових робіт', 3, { extraWorkDescription, extraWorkAmount, comment: extraWorkComment });
    createNotificationDraft(selectedOrder, 'Потрібне погодження додаткових робіт', 'Telegram', text, 'Створено');
    setNotificationPreview({ event: 'Потрібне погодження додаткових робіт', text, channel: 'Telegram' });
  };
  const copyNotificationPreview = async () => {
    if (!notificationPreview) return;
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(notificationPreview.text);
      createNotificationDraft(selectedOrder, notificationPreview.event, notificationPreview.channel, notificationPreview.text, 'Скопійовано');
      return;
    }
    window.prompt('Скопируйте текст уведомления', notificationPreview.text);
  };
  const confirmNotificationPreview = (channel: NotificationChannel, openUrl?: string) => {
    if (!notificationPreview) return;
    runWithWarnings('Менеджер отправил уведомление клиенту', () => {
      if (notificationPreview.event === 'Потрібне погодження додаткових робіт') {
        sendRepairApproval(selectedOrder, channel, { description: extraWorkDescription, amount: extraWorkAmount, comment: extraWorkComment });
      } else {
        sendClientNotification(selectedOrder, notificationPreview.event, channel, true, notificationPreview.text);
      }
      if (openUrl) window.open(openUrl, '_blank', 'noopener,noreferrer');
      setNotificationPreview(null);
    });
  };
  const renderWorkRowsHtml = () => (
    selectedOrder.works.length > 0
      ? selectedOrder.works.map((work, index) => `<tr><td class="center">${index + 1}</td><td>${escapeHtml(work.name)}</td><td class="center">${work.qty ?? 1}</td><td class="right">${escapeHtml(money(work.price))}</td><td class="right">${escapeHtml(money(work.price * (work.qty ?? 1)))}</td></tr>`).join('')
      : '<tr><td class="center">1</td><td>Роботи згідно замовлення</td><td class="center">1</td><td class="right">0 грн</td><td class="right">0 грн</td></tr>'
  );
  const renderPartRowsHtml = () => (
    selectedOrder.parts.length > 0
      ? selectedOrder.parts.map((part, index) => `<tr><td class="center">${index + 1}</td><td>${escapeHtml(products.find((item) => item.id === part.productId)?.name ?? part.productId)}</td><td class="center">${part.qty}</td><td class="right">${escapeHtml(money(part.price))}</td><td class="right">${escapeHtml(money(part.price * part.qty))}</td></tr>`).join('')
      : '<tr><td class="center">1</td><td>Запчастини не використовувались</td><td class="center">0</td><td class="right">0 грн</td><td class="right">0 грн</td></tr>'
  );
  const renderSimpleListHtml = (itemsHtml: string) => `<div>${itemsHtml || 'Не заповнено'}</div>`;
  const documentDateForKind = (kind: 'Рахунок на оплату' | 'Акт надання послуг' | 'Видаткова накладна' | 'Наряд на ремонт' | 'Гарантійний талон') => {
    if (kind === 'Наряд на ремонт') return extractDayKey(selectedOrder.intakeDate);
    if (kind === 'Акт надання послуг') return extractDayKey(selectedOrder.statusChangedAt);
    return today;
  };
  const orderDocumentHtml = useMemo(() => {
    const documentSnapshot = parseDocumentSnapshot(documents.find((document) => document.entityType === 'service_order' && document.entityId === selectedOrder.id && document.kind === selectedDocumentKind)?.snapshotData);
    const worksRows = renderWorkRowsHtml();
    const partsRows = renderPartRowsHtml();
    const worksList = selectedOrder.works.map((work) => `${escapeHtml(work.name)} (${work.qty ?? 1} шт.)`).join('<br />');
    const partsList = selectedOrder.parts.map((part) => `${escapeHtml(products.find((item) => item.id === part.productId)?.name ?? part.productId)} (${part.qty} шт.)`).join('<br />');
    const common = {
      date: escapeHtml(documentSnapshot?.date ?? documentDateForKind(selectedDocumentKind)),
      orderNumber: escapeHtml(documentSnapshot?.orderNumber ?? documentNumberByKind[selectedDocumentKind]),
      clientName: escapeHtml(documentSnapshot?.clientName ?? selectedOrder.client),
      phone: escapeHtml(documentSnapshot?.phone ?? selectedOrder.phone),
      device: escapeHtml(documentSnapshot?.device ?? selectedOrder.device),
      serialNumber: escapeHtml(documentSnapshot?.serialNumber ?? selectedOrder.serial),
      problem: escapeHtml(documentSnapshot?.problem ?? selectedOrder.issue),
      total: escapeHtml(documentSnapshot?.total ?? money(totals.total)),
      paid: escapeHtml(documentSnapshot?.paid ?? money(totals.paid)),
      debt: escapeHtml(documentSnapshot?.debt ?? money(totals.debt)),
      managerName: escapeHtml(documentSnapshot?.managerName ?? selectedOrder.manager),
      engineerName: escapeHtml(documentSnapshot?.engineerName ?? selectedOrder.engineer),
      clientTaxId: escapeHtml(documentSnapshot?.clientTaxId ?? orderClientDetails.edrpou ?? '—'),
      amountNet: escapeHtml(documentSnapshot?.amountNet ?? money(invoiceTotals(totals.total, companySettings.vatEnabled).totalWithoutVAT)),
      amountVat: escapeHtml(documentSnapshot?.amountVat ?? money(invoiceTotals(totals.total, companySettings.vatEnabled).vat)),
      amountGross: escapeHtml(documentSnapshot?.amountGross ?? money(invoiceTotals(totals.total, companySettings.vatEnabled).totalWithVAT)),
      amountWords: escapeHtml(documentSnapshot?.amountWords ?? numberToWordsUA(invoiceTotals(totals.total, companySettings.vatEnabled).totalWithVAT)),
      vatWords: escapeHtml(documentSnapshot?.vatWords ?? numberToWordsUA(invoiceTotals(totals.total, companySettings.vatEnabled).vat)),
      companyName: escapeHtml(documentSnapshot?.companyName ?? companySettings.companyName),
      companyEdrpou: escapeHtml(documentSnapshot?.companyEdrpou ?? companySettings.edrpou),
      companyIban: escapeHtml(documentSnapshot?.companyIban ?? companySettings.iban),
      companyBank: escapeHtml(documentSnapshot?.companyBank ?? companySettings.bank),
      companyMfo: escapeHtml(documentSnapshot?.companyMfo ?? companySettings.mfo),
      companyAddress: escapeHtml(documentSnapshot?.companyAddress ?? companySettings.address),
      companyPhone: escapeHtml(documentSnapshot?.companyPhone ?? companySettings.phone),
      companyIpn: escapeHtml(documentSnapshot?.companyIpn ?? companySettings.ipn),
      companyVatCertificate: escapeHtml(documentSnapshot?.companyVatCertificate ?? companySettings.vatCertificate),
    };
    const frozenWorksList = documentSnapshot?.worksList ?? worksList;
    const frozenPartsList = documentSnapshot?.partsList ?? partsList;
    if (selectedDocumentKind === 'Рахунок на оплату') {
      return renderTemplate(invoiceTemplateHtml, {
        'document.css': serviceDocumentPreviewCss,
        'invoice.number': common.orderNumber,
        'invoice.date': common.date,
        'invoice.managerName': common.managerName,
        'invoice.reason': escapeHtml(documentSnapshot?.invoiceReason ?? `Замовлення ${selectedOrder.id} від ${extractDayKey(selectedOrder.intakeDate)}`),
        'company.name': common.companyName,
        'company.edrpou': common.companyEdrpou,
        'company.iban': common.companyIban,
        'company.bank': common.companyBank,
        'company.mfo': common.companyMfo,
        'company.address': common.companyAddress,
        'company.phone': common.companyPhone,
        'company.email': escapeHtml(companySettings.mainEmail),
        'company.ipn': common.companyIpn,
        'company.vatCertificate': escapeHtml(`${documentSnapshot?.companyVatStatus ?? 'Платник ПДВ'} · ${documentSnapshot?.companyVatCertificate ?? companySettings.vatCertificate}`),
        'client.name': common.clientName,
        'client.phone': common.phone,
        'client.edrpou': common.clientTaxId,
        'items.rows': documentSnapshot?.invoiceItemsRows ?? `${worksRows}${partsRows}`,
        'totals.net': common.amountNet,
        'totals.vat': common.amountVat,
        'totals.gross': common.amountGross,
        'totals.paid': common.paid,
        'totals.debt': common.debt,
        'totals.words': escapeHtml(documentSnapshot?.totalWithVATWords ?? common.amountWords),
        'totals.vatWords': escapeHtml(documentSnapshot?.vatAmountWords ?? common.vatWords),
        'order.device': common.device,
        'order.serial': common.serialNumber,
        'order.problem': common.problem,
      });
    }
    if (selectedDocumentKind === 'Акт надання послуг') {
      return renderTemplate(actTemplateHtml, {
        'document.css': serviceDocumentPreviewCss,
        'act.number': common.orderNumber,
        'act.date': common.date,
        'act.managerName': common.managerName,
        'invoice.reference': escapeHtml(documentSnapshot?.linkedInvoiceNumber ? `Рахунок на оплату № ${documentSnapshot.linkedInvoiceNumber} від ${documentSnapshot.linkedInvoiceDate ?? common.date}` : 'Рахунок не створено'),
        'company.name': common.companyName,
        'company.address': common.companyAddress,
        'company.phone': common.companyPhone,
        'company.email': escapeHtml(companySettings.mainEmail),
        'company.vatCertificate': escapeHtml(`${documentSnapshot?.companyVatStatus ?? 'Платник ПДВ'} · ${documentSnapshot?.companyVatCertificate ?? companySettings.vatCertificate}`),
        'client.name': common.clientName,
        'client.phone': common.phone,
        'client.edrpou': common.clientTaxId,
        'order.id': escapeHtml(selectedOrder.id),
        'order.device': common.device,
        'order.problem': common.problem,
        'services.rows': documentSnapshot?.serviceItemsRows ?? renderWorkRowsHtml(),
        'products.rows': documentSnapshot?.productItemsRows ?? renderPartRowsHtml(),
        'totals.net': common.amountNet,
        'totals.vat': common.amountVat,
        'totals.gross': common.amountGross,
        'totals.words': escapeHtml(documentSnapshot?.totalWithVATWords ?? common.amountWords),
        'totals.vatWords': escapeHtml(documentSnapshot?.vatAmountWords ?? common.vatWords),
      });
    }
    if (selectedDocumentKind === 'Видаткова накладна') {
      return renderTemplate(waybillTemplateHtml, {
        'document.css': serviceDocumentPreviewCss,
        'waybill.number': common.orderNumber,
        'waybill.date': common.date,
        'waybill.managerName': common.managerName,
        'company.name': common.companyName,
        'company.edrpou': common.companyEdrpou,
        'company.address': common.companyAddress,
        'company.phone': common.companyPhone,
        'client.name': common.clientName,
        'client.edrpou': common.clientTaxId,
        'products.rows': documentSnapshot?.productItemsRows ?? renderPartRowsHtml(),
        'totals.net': common.amountNet,
        'totals.vat': common.amountVat,
        'totals.gross': common.amountGross,
        'totals.words': escapeHtml(documentSnapshot?.totalWithVATWords ?? common.amountWords),
        'totals.vatWords': escapeHtml(documentSnapshot?.vatAmountWords ?? common.vatWords),
      });
    }
    if (selectedDocumentKind === 'Гарантійний талон') {
      return `
        <!doctype html><html lang="uk"><head><meta charset="UTF-8" /><style>${serviceDocumentPreviewCss}</style><title>Гарантійний талон</title></head><body>
        <section class="sheet">
          <div class="company-block" style="margin-left:18mm;margin-top:4mm;"><div class="company-name">${escapeHtml(companyRequisites.name)}</div><div>${escapeHtml(companyRequisites.address)}</div><div>${escapeHtml(companyRequisites.phone)}</div><div>${escapeHtml(companySettings.mainEmail)}</div></div>
          <div class="title" style="margin-top:20mm;">ГАРАНТІЙНИЙ ТАЛОН</div><div class="subtitle">Дата продажу «${escapeHtml(common.date)}»</div>
          <table class="doc-table" style="width:165mm;margin-left:auto;margin-right:auto;"><thead><tr><th style="width:52mm;text-align:left;">Серійний номер</th><th style="text-align:left;">Найменування обладнання</th><th style="width:42mm;text-align:left;">Термін гарантії (міс.)</th></tr></thead><tbody><tr><td>${common.serialNumber}</td><td>${common.device}</td><td>6 місяців</td></tr></tbody></table>
          <div class="details-grid" style="width:165mm;margin:0 auto;"><div><strong>Клієнт:</strong> ${common.clientName}</div><div><strong>Телефон:</strong> ${common.phone}</div><div><strong>Замовлення:</strong> ${escapeHtml(selectedOrder.id)}</div><div><strong>Дата:</strong> ${common.date}</div></div>
          <div class="spacer-xl"></div><div class="center">Підпис ____________________________</div>
          <div class="spacer-xl"></div>
          <div class="justify" style="width:165mm;margin:0 auto;"><p>Цим документом сервісний центр гарантує, що в разі виходу з ладу обладнання протягом терміну гарантії, воно буде безоплатно відремонтовано або замінене на справне.</p><p>Гарантійний ремонт здійснюється за наявності гарантійного талону з пред'явленням дати продажу, терміну гарантії, найменування і серійного номера обладнання.</p><p>Гарантія не поширюється на витратні матеріали та збереження інформації на носіях даних.</p></div>
          <div class="meta-row" style="grid-template-columns:78mm 1fr;width:165mm;margin:12mm auto 0;"><div class="label">З умовами гарантійних зобов'язань ознайомлений</div><div class="value-line">${common.clientName}</div></div>
        </section></body></html>`;
    }
    return `
      <!doctype html><html lang="uk"><head><meta charset="UTF-8" /><style>${serviceDocumentPreviewCss}</style><title>Наряд</title></head><body>
      <section class="sheet compact">
        <div class="small">Сервіс-центр: ${escapeHtml(companyRequisites.name)}, код ${escapeHtml(companyRequisites.edrpou)}, ${escapeHtml(companyRequisites.address)}, ${escapeHtml(companyRequisites.phone)}, ${escapeHtml(companySettings.mainEmail)}</div>
        <div class="doc-line">Наряд № ${common.orderNumber} від ${common.date}</div>
        <div class="meta-grid">
          <div class="meta-row"><div class="label">Замовник:</div><div class="value-line">${common.clientName}</div></div>
          <div class="meta-row"><div class="label">Адреса:</div><div class="value-line">${escapeHtml(orderClientDetails.address)}</div></div>
          <div class="meta-row"><div class="label">Телефон:</div><div class="value-line">${common.phone}</div></div>
          <div class="meta-row"><div class="label">E-mail:</div><div class="value-line">________________</div></div>
          <div class="meta-row"><div class="label">Представник:</div><div class="value-line">${common.clientName}</div></div>
          <div class="meta-row"><div class="label">Гарантійний:</div><div class="value-line">НІ</div></div>
        </div>
        <p><strong>Сервісний центр прийняв виріб:</strong></p>
        <table class="doc-table"><tbody>
          <tr><td style="width:45mm;">Найменування виробу</td><td colspan="3">${common.device}</td></tr>
          <tr><td>Дата продажу згідно гарантійного талону</td><td colspan="3">________________</td></tr>
          <tr><td>Серійний номер</td><td>${common.serialNumber}</td><td style="width:28mm;">Продукт номер</td><td>________________</td></tr>
          <tr><td>Комплектність виробу</td><td colspan="3">________________</td></tr>
          <tr><td>Стан виробу</td><td colspan="3">________________</td></tr>
          <tr><td>Заявлена несправність</td><td colspan="3">${common.problem}</td></tr>
          <tr><td>Виконані роботи</td><td colspan="3">${frozenWorksList || '________________'}</td></tr>
          <tr><td>Використані запчастини</td><td colspan="3">${frozenPartsList || '________________'}</td></tr>
          <tr><td>Примітки</td><td colspan="3">________________</td></tr>
        </tbody></table>
        <div class="details-grid"><div><strong>Сума робіт:</strong> ${common.total}</div><div><strong>Оплачено:</strong> ${common.paid}</div><div><strong>Борг:</strong> ${common.debt}</div><div><strong>Інженер:</strong> ${common.engineerName}</div></div>
        <div class="terms-box tiny"><p><strong>Умови обслуговування в сервісному центрі НВКПП "Спектр-АС"</strong></p><p>Виріб вважається гарантійним тільки після попередньої діагностики. Виконуємо ремонт тільки заявлених несправностей. Сервісний центр не несе відповідальності за програмне забезпечення та збереження інформації.</p><p>Про стан виконання замовлення можна дізнатися за телефонами сервісного центру. При отриманні виробу необхідно назвати номер квитанції або замовлення.</p></div>
        <div class="signature-row"><div class="signature-block"><div>Виконав ремонт:</div><div class="signature-line"></div><div class="right">${common.engineerName}</div></div><div class="signature-block"><div>Менеджер</div><div class="signature-line"></div><div class="right">${common.managerName}</div></div><div class="signature-block"><div>Дата завершення ремонту</div><div class="signature-line">${escapeHtml(extractDayKey(selectedOrder.statusChangedAt))}</div></div></div>
        <div class="meta-row"><div class="label">З умовами ремонту згоден:</div><div class="value-line">${common.clientName}</div></div>
      </section></body></html>`;
  }, [documents, orderClientDetails.address, products, selectedDocumentKind, selectedOrder, totals.debt, totals.paid, totals.total]);
  const openDocumentPreview = (kind: 'Рахунок на оплату' | 'Акт надання послуг' | 'Видаткова накладна' | 'Наряд на ремонт' | 'Гарантійний талон') => {
    const validationError = validateDocumentPreview(kind);
    if (validationError) {
      notifyUser(validationError);
      return;
    }
    setSelectedDocumentKind(kind);
    setDocumentPreviewNonce((value) => value + 1);
    setIsDocumentPreviewOpen(true);
  };
  const printPreviewDocument = () => {
    const frameWindow = documentIframeRef.current?.contentWindow;
    if (!frameWindow) return;
    if (selectedDocumentKind === 'Рахунок на оплату' || selectedDocumentKind === 'Акт надання послуг' || selectedDocumentKind === 'Видаткова накладна') {
      if (!hasDocument(selectedDocumentKind)) {
        notifyUser(`Спочатку створіть документ "${selectedDocumentKind}".`);
        return;
      }
      printServiceOrderDocument(selectedDocumentKind, selectedOrder);
    } else if (selectedDocumentKind === 'Наряд на ремонт' || selectedDocumentKind === 'Гарантійний талон') {
      ensureOrderDocumentRecord(selectedDocumentKind, selectedOrder);
      logOrderDocumentPrint(selectedOrder, selectedDocumentKind);
    }
    if (activeDocumentSuggestion?.kind === selectedDocumentKind) clearSuggestedDocumentAction();
    frameWindow.focus();
    frameWindow.print();
  };
  const validateDocumentPreview = (kind: 'Рахунок на оплату' | 'Акт надання послуг' | 'Видаткова накладна' | 'Наряд на ремонт' | 'Гарантійний талон') => {
    if (kind === 'Наряд на ремонт') return '';
    if (['Рахунок на оплату', 'Акт надання послуг', 'Видаткова накладна'].includes(kind) && !hasDocument(kind as DocumentKind)) return `Спочатку створіть документ "${kind}".`;
    if (kind === 'Видаткова накладна' && selectedOrder.parts.filter((part) => part.status !== 'Повернення').length === 0) return 'Видаткова накладна доступна тільки якщо в замовленні є товарні позиції.';
    if (kind === 'Рахунок на оплату' && totals.total <= 0) return 'Рахунок доступний тільки коли в замовленні є сума.';
    if (kind === 'Акт надання послуг' && selectedOrder.status !== 'Готовий до видачі') return 'Акт виконаних робіт доступний тільки на етапі "Готово".';
    if (kind === 'Гарантійний талон' && !['Видано', 'Закрито'].includes(selectedOrder.status)) return 'Гарантійний талон доступний тільки після видачі клієнту.';
    return '';
  };
  const openLabelPreview = () => {
    setIsLabelPreviewOpen(true);
  };
  const printLabelPreview = () => {
    const labelWindow = window.open('', '_blank', 'width=420,height=620');
    if (!labelWindow) return;
    const shortDevice = selectedOrder.device.length > 48 ? `${selectedOrder.device.slice(0, 48)}...` : selectedOrder.device;
    const orderAgeDays = daysSince(selectedOrder.intakeDate);
    labelWindow.document.write(`
      <!doctype html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8" />
          <title>Наклейка ${escapeHtml(selectedOrder.id)}</title>
          <style>
            @page { size: 100mm 150mm; margin: 6mm; }
            body { font-family: Arial, sans-serif; margin: 0; color: #0f172a; }
            .label { border: 2px solid #0f172a; padding: 8mm 7mm; }
            .top { display: flex; justify-content: space-between; align-items: center; gap: 8mm; }
            .order { font-size: 26px; font-weight: 700; }
            .cell { margin-top: 5mm; font-size: 30px; font-weight: 800; line-height: 1.05; letter-spacing: 0.5px; }
            .status { padding: 6px 10px; font-size: 12px; font-weight: 700; color: #fff; background: ${labelStatusColor}; }
            .qr { width: 34mm; height: 34mm; border: 1px solid #94a3b8; object-fit: contain; }
            .row { margin-top: 6mm; }
            .label-title { font-size: 11px; text-transform: uppercase; color: #475569; margin-bottom: 2px; }
            .value { font-size: 15px; font-weight: 600; }
            .state { display: inline-block; margin-top: 4mm; padding: 4px 8px; border: 1px solid #0f172a; font-size: 12px; font-weight: 700; }
          </style>
        </head>
        <body>
          <div class="label">
            <div class="top">
              <div>
                <div class="order">${escapeHtml(selectedOrder.id)}</div>
                <div class="cell">${escapeHtml(selectedOrder.locationCode ?? 'БЕЗ ЯЧЕЙКИ')}</div>
                <div class="status">${escapeHtml(labelStatusText)}</div>
                <div class="state">${escapeHtml(labelLocationStateText)}</div>
              </div>
              <img class="qr" src="https://api.qrserver.com/v1/create-qr-code/?size=180x180&data=${encodeURIComponent(directOrderUrl)}" alt="QR" />
            </div>
            <div class="row"><div class="label-title">Дата прийому</div><div class="value">${escapeHtml(extractDayKey(selectedOrder.intakeDate))}</div></div>
            <div class="row"><div class="label-title">Вік замовлення</div><div class="value">${escapeHtml(String(orderAgeDays))} дн.</div></div>
            <div class="row"><div class="label-title">Клієнт</div><div class="value">${escapeHtml(selectedOrder.client)}</div></div>
            <div class="row"><div class="label-title">Пристрій</div><div class="value">${escapeHtml(shortDevice)}</div></div>
          </div>
        </body>
      </html>
    `);
    labelWindow.document.close();
    labelWindow.focus();
    labelWindow.print();
  };
  const workCompensationRows = selectedOrder.works.map((work, index) => {
    const employee = work.engineer ?? selectedOrder.engineer;
    const rule = payrollRules.find((item) => item.employee === employee);
    const qty = work.qty ?? 1;
    return {
      id: `${selectedOrder.id}-work-${index}`,
      workType: workTypeLabel(work),
      employee,
      qty,
      rate: workRateLabel(work, rule),
      workAmount: orderWorkAmount(work),
      earning: payrollForWork(selectedOrder, work, rule),
      workName: work.name,
    };
  });
  const totalEmployeeEarning = workCompensationRows.reduce((sum, row) => sum + row.earning, 0);
  const visibleLifecycle = isEngineerRole ? engineerLifecycleDisplay : canManageClientWork ? managerLifecycleDisplay : orderLifecycleDisplay;
  const visibleNextStatuses = nextStatuses.filter((status) => (
    isEngineerRole
      ? engineerStatuses.includes(status)
      : canManageClientWork
        ? managerStatuses.includes(status)
        : true
  ));
  const roleFocusTitle = isManagerRole ? 'Клиент и выдача' : isSupervisorRole ? 'Зона контроля' : 'Рабочая зона';
  const roleFocusItems = isEngineerRole
    ? [
          { label: 'Текущий этап', value: statusStage(selectedOrder.status) },
          { label: 'Запчасти', value: `${selectedOrder.parts.length}` },
          { label: 'Тест', value: selectedOrder.testResult ? 'Заполнен' : 'Ожидается' },
          { label: 'Готовность', value: isReadyForClient ? 'Да' : 'В работе' },
      ]
    : isManagerRole
      ? [
          { label: 'Клиент', value: selectedOrder.client },
          { label: 'Текущий этап', value: statusStage(selectedOrder.status) },
          { label: 'Оплата', value: simplePaymentStatus },
          { label: 'Выдача', value: canIssueOrder ? 'Можно выдать' : 'Пока рано' },
        ]
      : [
          { label: 'Статус', value: statusStage(selectedOrder.status) },
          { label: 'Финансы', value: money(totals.total) },
          { label: 'Долг', value: money(totals.debt) },
          { label: 'Контроль', value: `${statusAge} дн. в статусе` },
        ];
  useEffect(() => {
    setPendingLocationCode(selectedOrder.locationCode ?? '');
    setIsMiniCashOpen(false);
  }, [selectedOrder.id, selectedOrder.locationCode]);

  useEffect(() => {
    if (!activeDocumentSuggestion?.autoOpen) return;
    if (!canPreviewOrderDocuments) return;
    const validationError = validateDocumentPreview(activeDocumentSuggestion.kind);
    if (validationError) {
      notifyUser(validationError);
      clearSuggestedDocumentAction();
      return;
    }
    setSelectedDocumentKind(activeDocumentSuggestion.kind);
    setDocumentPreviewNonce((value) => value + 1);
    setIsDocumentPreviewOpen(true);
    clearSuggestedDocumentAction();
  }, [activeDocumentSuggestion, canPreviewOrderDocuments, clearSuggestedDocumentAction, notifyUser, selectedOrder.id]);
  useEffect(() => {
    if (!activeDashboardFocus || !canShowManagerView) return;
    if (activeDashboardFocus.target === 'payment') {
      const paymentInput = document.querySelector<HTMLInputElement>('input[aria-label="Сума оплати"]');
      paymentInput?.scrollIntoView({ behavior: 'smooth', block: 'center' });
      paymentInput?.focus();
      notifyUser(`Відкрито ${selectedOrder.id}: наступний крок — прийняти оплату.`);
    }
    if (activeDashboardFocus.target === 'issue') {
      const processPanel = document.querySelector<HTMLElement>('.process-panel');
      processPanel?.scrollIntoView({ behavior: 'smooth', block: 'start' });
      if (canPreviewAct) openDocumentPreview('Акт надання послуг');
      notifyUser(`Відкрито ${selectedOrder.id}: наступний крок — видати замовлення.`);
    }
    if (activeDashboardFocus.target === 'approval') {
      const approvalPanel = document.querySelector<HTMLElement>('.approval-panel');
      approvalPanel?.scrollIntoView({ behavior: 'smooth', block: 'start' });
      notifyUser(`Відкрито ${selectedOrder.id}: наступний крок — отримати відповідь клієнта.`);
    }
    clearDashboardFocus();
  }, [activeDashboardFocus, canPreviewAct, canShowManagerView, clearDashboardFocus, notifyUser, selectedOrder.id]);
  return (
    <section className="panel">
      <div className="panel-heading">
        <h2>{selectedOrder.id} · {selectedOrder.device}</h2>
        <span className={statusClass[selectedOrder.status]}>{statusDisplay(selectedOrder.status)}</span>
      </div>
      <div className="details-grid">
        {canShowManagerView && <Info label="Клиент" value={selectedOrder.client} />}
        {canShowManagerView && <Info label="Телефон" value={selectedOrder.phone} />}
        {canShowManagerView && <Info label="Устройство" value={selectedOrder.device} />}
        {canShowManagerView && <Info label="Проблема" value={selectedOrder.issue} />}
        {canShowManagerView && <Info label="Комплектація" value={selectedOrder.appearance ?? 'Не вказано'} />}
        {canShowManagerView && <Info label="Коментар" value={selectedOrder.intakeComment ?? 'Без коментаря'} />}
        {canShowManagerView && <Info label="Текущий этап" value={statusStage(selectedOrder.status)} />}
        {canShowManagerView && <Info label="Сума" value={displayedOrderAmount} />}
        {canShowManagerView && <Info label="Оплачено" value={money(totals.paid)} />}
        {canShowManagerView && <Info label="Долг" value={money(totals.debt)} />}
        {canShowManagerView && <Info label="Статус оплаты" value={simplePaymentStatus} />}
        {canShowManagerView && <Info label="Клиент уведомлён" value={clientNotified ? 'Уведомлён' : 'Не уведомлён'} />}
        {canShowManagerView && <Info label="Останнє повідомлення" value={selectedOrder.lastNotificationType ? `${notificationDisplay(selectedOrder.lastNotificationType)} · ${selectedOrder.lastNotificationAt ?? 'без дати'}` : 'Ще не надсилали'} />}
        {canShowManagerView && <Info label="К выдаче" value={canIssueOrder ? 'Можно выдать' : 'Ещё не готово'} />}
        {canShowManagerView && <Info label="Локація" value={selectedOrder.locationCode ?? 'Не призначено'} />}
        {canShowManagerView && <Info label="Інженер" value={selectedOrder.engineer} />}
        {canShowManagerView && <Info label="Створив / прийняв" value={orderCreatorName} />}
        {canShowManagerView && <Info label="Дата прийому" value={selectedOrder.intakeDate} />}
        {canShowManagerView && <Info label="Вік замовлення" value={`${totalAge} дн.`} />}
        {canShowManagerView && <Info label="Терміновість" value={selectedOrder.urgent ? 'ТЕРМІНОВО' : 'Звичайний пріоритет'} />}

        {canShowEngineerView && <Info label="Проблема" value={selectedOrder.issue} />}
        {canShowEngineerView && <Info label="Текущий этап" value={statusStage(selectedOrder.status)} />}
        {canShowEngineerView && <Info label="Запчасти" value={String(selectedOrder.parts.length)} />}
        {canShowEngineerView && <Info label="Текущий статус" value={`${statusDisplay(selectedOrder.status)} · ${statusAge} дн.`} />}
        {canShowEngineerView && <Info label="Ячейка" value={selectedOrder.locationCode ?? 'не призначено'} />}
        {canShowEngineerView && <Info label="Стан ячейки" value={selectedOrder.locationStatus ?? 'не задано'} />}
        {canShowEngineerView && <Info label="Таймер у інженера" value={engineerElapsedLabel(selectedOrder.takenInRepairAt ?? selectedOrder.statusChangedAt)} />}

        {canShowReadonlyOverview && <Info label="Сумма ремонта" value={money(totals.total)} />}
        {canShowReadonlyOverview && <Info label="Оплачено" value={money(totals.paid)} />}
        {canShowReadonlyOverview && <Info label="Долг" value={money(totals.debt)} />}
        {canShowReadonlyOverview && <Info label="Статус оплаты" value={simplePaymentStatus} />}
        {canShowReadonlyOverview && <Info label="Локація" value={selectedOrder.locationCode ?? 'Не призначено'} />}

        {!canShowEngineerView && !canShowManagerView && !canShowReadonlyOverview && <Info label="Клієнт" value={`${selectedOrder.client}, ${selectedOrder.phone}`} />}
        {!canShowEngineerView && !canShowManagerView && !canShowReadonlyOverview && <Info label="Серійний номер" value={selectedOrder.serial} />}
        {!canShowEngineerView && !canShowManagerView && !canShowReadonlyOverview && <Info label="Інженер" value={selectedOrder.engineer} />}
        {!canShowEngineerView && !canShowManagerView && !canShowReadonlyOverview && <Info label="Підтвердження інженера" value={selectedOrder.engineerAcceptedAt ?? 'Ще не взяв у роботу'} />}
        {!canShowEngineerView && !canShowManagerView && !canShowReadonlyOverview && <Info label="Менеджер" value={selectedOrder.manager} />}
        {!canShowEngineerView && !canShowManagerView && !canShowReadonlyOverview && <Info label="Несправність" value={selectedOrder.issue} />}
        {!canShowEngineerView && !canShowManagerView && !canShowReadonlyOverview && canSeeFinance && <Info label="Загальна сума" value={money(totals.total)} />}
        {!canShowEngineerView && !canShowManagerView && !canShowReadonlyOverview && canSeeFinance && <Info label="Сума робіт" value={money(totals.works)} />}
        {!canShowEngineerView && !canShowManagerView && !canShowReadonlyOverview && canSeeFinance && <Info label="Сума запчастин" value={money(totals.parts)} />}
        {!canShowEngineerView && !canShowManagerView && !canShowReadonlyOverview && canSeeFinance && <Info label="Доставка" value={money(totals.delivery)} />}
        {!canShowEngineerView && !canShowManagerView && !canShowReadonlyOverview && canSeeFinance && <Info label="Прочі витрати" value={money(totals.additionalExpenses + totals.consumables + totals.externalServices)} />}
        {!canShowEngineerView && !canShowManagerView && !canShowReadonlyOverview && canSeeFinance && <Info label="Зарплата інженера" value={showCost ? money(totals.engineerSalary) : 'Приховано правами доступу'} />}
        {!canShowEngineerView && !canShowManagerView && !canShowReadonlyOverview && canSeeFinance && <Info label="Закупка запчастин" value={showCost ? money(totals.purchaseCost) : 'Приховано правами доступу'} />}
        {!canShowEngineerView && !canShowManagerView && !canShowReadonlyOverview && canSeeFinance && <Info label="Підсумковий прибуток" value={showCost ? `${money(totals.profit)}${totals.isFinal ? '' : ' · фінально після закриття'}` : 'Приховано правами доступу'} />}
        {!canShowEngineerView && !canShowManagerView && !canShowReadonlyOverview && canSeeFinance && <Info label="Оплата" value={`${simplePaymentStatus} · оплачено ${money(totals.paid)} · борг ${money(totals.debt)}`} />}
        {!canShowEngineerView && !canShowManagerView && !canShowReadonlyOverview && canSeeFinance && <Info label="Собівартість" value={showCost ? money(totals.cost) : 'Приховано правами доступу'} />}
        {!canShowEngineerView && !canShowManagerView && !canShowReadonlyOverview && <Info label="Контроль строків" value={`У статусі ${statusAge} дн. · з прийому ${totalAge} дн. · обіцяно ${selectedOrder.promisedDate}`} />}
      </div>
      <div className="linear-status">
        {visibleLifecycle.map((step) => (
          <span key={step.key} className={selectedOrder.status === step.key || statusStage(step.key) === statusStage(selectedOrder.status) ? 'linear-status-active' : ''}>{step.label}</span>
        ))}
      </div>
      {canShowManagerView && (
        <div className="manager-order-layout">
          <div className="manager-order-main">
            <section className="panel-subsection manager-center-section">
              <div className="panel-heading">
                <h2>Робочий центр замовлення</h2>
                <span>{selectedOrder.id} · {statusStage(selectedOrder.status)}</span>
              </div>
              <div className="details-grid manager-center-grid">
                <Info label="Номер замовлення" value={selectedOrder.id} />
                <Info label="Статус" value={statusDisplay(selectedOrder.status)} />
                <Info label="Комірка" value={selectedOrder.locationCode ?? 'Не призначено'} />
                <Info label="Клієнт" value={selectedOrder.client} />
                <Info label="Телефон" value={selectedOrder.phone} />
                <Info label="Пристрій" value={selectedOrder.device} />
                <Info label="Серійний номер" value={selectedOrder.serial} />
                <Info label="Проблема" value={selectedOrder.issue} />
                <Info label="Комплектація" value={selectedOrder.appearance ?? 'Не вказано'} />
                <Info label="Коментар" value={selectedOrder.intakeComment ?? 'Без коментаря'} />
                <Info label="Інженер" value={selectedOrder.engineer} />
                <Info label="Створив / прийняв" value={orderCreatorName} />
                <Info label="Дата прийому" value={selectedOrder.intakeDate} />
                <Info label="Вік замовлення" value={`${totalAge} дн.`} />
                <Info label="Сума" value={displayedOrderAmount} />
                <Info label="Оплачено" value={money(totals.paid)} />
                <Info label="Борг" value={money(totals.debt)} />
                <Info label="Наступний крок" value={managerNextStep} />
                <Info label="Стан угоди" value={dealClosed ? 'Сделка закрыта' : 'Ще в роботі'} />
                <Info label="Статус акта" value={actDocumentStatus} />
                <Info label="ПН" value={selectedOrderTaxInvoice ? `${selectedOrderTaxInvoice.number} · ${selectedOrderTaxInvoice.status}` : (selectedOrder.legalEntity || selectedOrder.vatStatus ? 'Не створено' : 'Не потрібна')} />
                <Info label="Стан комірки" value={selectedOrder.locationStatus ?? 'Не задано'} />
              </div>
              {documentsOutdated && (
                <p style={{ color: '#b45309', fontWeight: 700, marginTop: '14px' }}>
                  {orderDocumentsMismatchText(selectedOrder)} Потрібно пересоздати нові документи.
                </p>
              )}
              {selectedOrder.returnedToCellAt && (
                <p style={{ color: '#166534', fontWeight: 700, marginTop: '14px' }}>
                  Готово до видачі: в комірці {selectedOrder.locationCode}
                </p>
              )}
              <div className="manager-order-tax-callout">
                <div className="manager-order-tax-callout-head">
                  <div>
                    <strong>Податкова накладна</strong>
                    <span>{taxInvoiceHint}</span>
                  </div>
                  <span className={`manager-order-tax-status ${taxInvoiceBadgeClass}`}>
                    {taxInvoiceBadgeText}
                  </span>
                </div>
                <div className="manager-order-tax-callout-actions">
                  {isVatOrder && !selectedOrderTaxInvoice && (
                    <button type="button" className="manager-order-tax-primary" onClick={() => createTaxInvoiceForOrder(selectedOrder.id)}>
                      Створити ПН
                    </button>
                  )}
                  {isVatOrder && selectedOrderTaxInvoice && selectedOrderTaxInvoice.status !== 'Зареєстровано' && (
                    <button type="button" className="manager-order-tax-primary" onClick={() => registerTaxInvoice(selectedOrder.id)}>
                      Зареєструвати ПН
                    </button>
                  )}
                </div>
              </div>
            </section>
            <section className="process-panel">
              <div className="panel-heading">
                <h2>Дії менеджера</h2>
                <span>1 кнопка = 1 етап</span>
              </div>
              <div className="simple-action-panel">
                <button type="button" onClick={managerAcceptIntoWork} disabled={!managerCanAcceptIntoWork}>Передати інженеру</button>
                <button type="button" onClick={() => sendRepairApproval(selectedOrder, 'Telegram')} disabled={statusActionCriticalMessages.length > 0}>Відправити на погодження</button>
                <button type="button" onClick={() => openLabelPreview()}>Печать наклейки</button>
                <button type="button" onClick={() => openDocumentPreview('Наряд на ремонт')} disabled={!canPreviewWorkOrder}>Печать наряда</button>
                <button type="button" onClick={() => canPreviewAct ? openDocumentPreview('Акт надання послуг') : canPreviewInvoice ? openDocumentPreview('Рахунок на оплату') : openDocumentPreview('Наряд на ремонт')}>Відкрити документи</button>
                <button type="button" onClick={managerMarkReadyForIssue} disabled={!managerCanMarkReadyForIssue}>Підготувати до видачі</button>
                <button type="button" onClick={() => setIsMiniCashOpen(true)} disabled={!managerCanTakePayment}>Прийняти оплату</button>
                <button type="button" onClick={handleManagerIssueClick} disabled={!managerCanIssueFromCard || issueCriticalMessages.length > 0}>Видати замовлення</button>
              </div>
              {isMiniCashOpen && (
                <div className="panel-subsection" style={{ marginTop: '12px' }}>
                  <div className="panel-heading">
                    <h2>Міні-каса</h2>
                    <span>оплата і видача в одному кроці</span>
                  </div>
                  <div className="details-grid">
                    <Info label="Сума" value={money(totals.total)} />
                    <Info label="Вже оплачено" value={money(totals.paid)} />
                    <Info label="До оплати" value={money(totals.debt)} />
                  </div>
                  <div className="simple-action-panel">
                    <button type="button" onClick={() => oneClickManagerIssue(selectedOrder, { method: 'Готівка' })} disabled={totals.debt <= 0}>Готівка</button>
                    <button type="button" onClick={() => oneClickManagerIssue(selectedOrder, { method: 'Картка' })} disabled={totals.debt <= 0}>Картка</button>
                    <button type="button" onClick={() => setIsMiniCashOpen((value) => !value)}>Частково</button>
                  </div>
                  <div className="action-row payment-actions">
                    <input aria-label="Сума оплати" type="number" min={1} max={totals.debt || undefined} placeholder={String(totals.debt)} value={paymentAmount || ''} onChange={(event) => setPaymentAmount(Number(event.target.value))} />
                    <select aria-label="Спосіб часткової оплати" value={paymentMethod} onChange={(event) => setPaymentMethod(event.target.value as PaymentMethod)}>
                      <option value="Готівка">Готівка</option>
                      <option value="Картка">Картка</option>
                    </select>
                    <button type="button" onClick={() => oneClickManagerIssue(selectedOrder, { method: paymentMethod, amount: effectivePaymentAmount, partial: true })} disabled={effectivePaymentAmount <= 0}>Прийняти частково</button>
                  </div>
                </div>
              )}
            </section>
          </div>
          <aside className="manager-order-side">
            {(controlState.critical.length > 0 || controlState.warnings.length > 0) && (
              <section className="panel-subsection">
                <div className="panel-heading">
                  <h2>Контроль помилок</h2>
                  <span>{controlState.critical.length} блокує · {controlState.warnings.length} попереджає</span>
                </div>
                {controlState.critical.length > 0 && (
                  <div className="task-list">
                    {controlState.critical.map((item) => <Task key={item.code} icon={<X />} title="Критична помилка" text={item.message} />)}
                  </div>
                )}
                {controlState.warnings.length > 0 && (
                  <div className="task-list">
                    {controlState.warnings.map((item) => <Task key={item.code} icon={<Bell />} title="Попередження" text={item.message} />)}
                  </div>
                )}
              </section>
            )}
            <section className="panel-subsection">
              <div className="panel-heading">
                <h2>Сповіщення клієнту</h2>
                <span>{clientNotified ? 'Клієнта повідомлено' : 'Ще не повідомлено'}</span>
              </div>
              <div className="task-list">
                {notifications.slice(0, 3).map((message) => (
                  <Task key={message.id} icon={<Bell />} title={`${notificationDisplay(message.event)} · ${message.channel}`} text={`${message.createdAt} · ${message.status}`} />
                ))}
                {notifications.length === 0 && <div className="empty-state">Історія сповіщень з’явиться тут після відправки клієнту.</div>}
              </div>
            </section>
            <section className="panel-subsection">
              <div className="panel-heading"><h2>Історія та логи</h2><span>{selectedOrder.statusHistory.length} записів</span></div>
              <div className="task-list">
                {selectedOrder.statusHistory.slice(0, 4).map((entry) => (
                  <Task key={entry.id} icon={<History />} title={`${entry.oldStatus ?? 'Старт'} -> ${entry.newStatus}`} text={`${entry.changedAt} · ${entry.changedBy}. ${entry.comment}`} />
                ))}
              </div>
            </section>
            <section className="panel-subsection">
              <div className="panel-heading"><h2>Версії замовлення</h2><span>{orderVersions.filter((item) => item.orderId === selectedOrder.id).length}</span></div>
              <div className="task-list">
                {orderVersions.filter((item) => item.orderId === selectedOrder.id).slice(0, 4).map((version) => (
                  <Task key={version.id} icon={<Archive />} title={`Версія ${version.versionNo}`} text={`${version.createdAt} · ${version.createdBy}. ${version.reason}`} />
                ))}
              </div>
            </section>
          </aside>
        </div>
      )}
      {!canShowManagerView && (controlState.critical.length > 0 || controlState.warnings.length > 0) && (
        <section className="panel-subsection">
          <div className="panel-heading">
            <h2>Контроль ошибок сотрудников</h2>
            <span>{controlState.critical.length} блокирует · {controlState.warnings.length} предупреждает</span>
          </div>
          {controlState.critical.length > 0 && (
            <div className="task-list">
              {controlState.critical.map((item) => <Task key={item.code} icon={<X />} title="Критическая ошибка" text={item.message} />)}
            </div>
          )}
          {controlState.warnings.length > 0 && (
            <div className="task-list">
              {controlState.warnings.map((item) => <Task key={item.code} icon={<Bell />} title="Предупреждение" text={item.message} />)}
            </div>
          )}
        </section>
      )}
      {!canShowReadonlyOverview && !canShowManagerView && !isEngineerRole && visibleNextStatuses.length > 0 && (
        <section className="panel-subsection">
          <div className="panel-heading">
            <h2>{isEngineerRole ? 'Следующие технические этапы' : 'Следующие клиентские этапы'}</h2>
            <span>показываются по текущему orderStatusFlow</span>
          </div>
          <div className="task-list">
            {visibleNextStatuses.map((status) => (
              <Task key={status} icon={<History />} title={transitionDisplay(status)} text={`Следующий статус: ${statusDisplay(status)}.`} />
            ))}
          </div>
        </section>
      )}
      {canShowRoleFocusBlock && <section className="panel-subsection">
        <div className="panel-heading">
          <h2>{roleFocusTitle}</h2>
          <span>{roleDisplay(activeUser.role)}</span>
        </div>
        <div className="details-grid">
          {roleFocusItems.map((item) => <Info key={item.label} label={item.label} value={item.value} />)}
        </div>
      </section>}
      {canShowQrPanel && <section className="qr-panel">
        <div>
          <h3>QR-код замовлення</h3>
          <p>Скан відкриває картку {selectedOrder.id} напряму. QR друкується на квитанції та може клеїтися на техніку.</p>
          <strong>{directOrderUrl}</strong>
        </div>
        <img
          alt={`QR ${selectedOrder.id}`}
          src={`https://api.qrserver.com/v1/create-qr-code/?size=132x132&data=${encodeURIComponent(directOrderUrl)}`}
        />
      </section>}
      {canSendApproval && <section className="panel-subsection approval-panel">
        <div className="panel-heading">
          <h2>Согласование доп. работ</h2>
          <span>{approval?.status ?? 'Не відправлено'}</span>
        </div>
        <div className="details-grid">
          <Info label="Сума для клієнта" value={approval ? money(approval.totalAmount) : money(totals.total)} />
          <Info label="Відправлено" value={approval ? `${approval.sentAt} · ${approval.sentChannel}` : 'Ще не відправлено'} />
          <Info label="Відповідь" value={approval?.responseAt ? `${approval.responseAt} · ${approval.responseIp}` : 'Немає відповіді'} />
          <Info label="Статус доп. робіт" value={selectedOrder.pendingExtraApproval ? 'Очікує погодження клієнта' : approval?.status ?? 'Немає активного погодження'} />
        </div>
        <div className="action-row payment-actions">
          <input aria-label="Опис додаткових робіт" value={extraWorkDescription} onChange={(event) => setExtraWorkDescription(event.target.value)} placeholder="Що саме потрібно зробити додатково" />
          <input aria-label="Сума додаткових робіт" type="number" min={0} value={extraWorkAmount || ''} onChange={(event) => setExtraWorkAmount(Number(event.target.value))} placeholder="Сума доп. робіт" />
          <input aria-label="Коментар менеджера" value={extraWorkComment} onChange={(event) => setExtraWorkComment(event.target.value)} placeholder="Коментар для клієнта" />
        </div>
        {approval && (
          <div className="approval-card">
            <div>
              <strong>Сторінка клієнта</strong>
              <p>{approval.clientUrl}</p>
            </div>
            <div>
              <strong>Зафіксовано в пропозиції</strong>
              <p>{approval.worksSnapshot.length} робіт · {approval.partsSnapshot.length} запчастин · {money(approval.totalAmount)}</p>
            </div>
          </div>
        )}
        <div className="action-row">
          <button type="button" onClick={openExtraApprovalPreview}>Согласовать доп. работы</button>
          <button type="button" onClick={() => approval && setNotificationPreview({ event: 'Потрібне погодження додаткових робіт', text: approval.messageText, channel: approval.sentChannel })} disabled={!approval}>Повторити повідомлення</button>
          <button type="button" onClick={() => approval && recordApprovalResponse(approval.id, true, true)} disabled={!approval}>Клієнт погодив</button>
          <button type="button" onClick={() => approval && markApprovalNoAnswer(approval.id)} disabled={!approval || approval.status !== 'Очікує відповідь'}>Немає відповіді</button>
          <button type="button" onClick={() => approval && recordApprovalResponse(approval.id, false, true)} disabled={!approval}>Клієнт відмовився</button>
        </div>
        {approval && (
          <div className="client-approval-preview">
            <h3>Сторінка для клієнта</h3>
            <p>{approval.device}</p>
            <p>{approval.issue}</p>
            {approval.extraWorkDescription && <p><strong>Додаткові роботи:</strong> {approval.extraWorkDescription}</p>}
            {approval.extraWorkAmount ? <p><strong>Сума доп. робіт:</strong> {money(approval.extraWorkAmount)}</p> : null}
            <div className="table approval-table">
              <div className="table-row table-head"><span>Позиція</span><span>Кількість</span><span>Сума</span></div>
              {approval.worksSnapshot.map((work) => <div className="table-row" key={work.name}><span>{work.name}<small>{work.serviceType}</small></span><span>{work.qty ?? 1}</span><span>{money(work.price * (work.qty ?? 1))}</span></div>)}
              {approval.partsSnapshot.map((part) => <div className="table-row" key={part.name}><span>{part.name}</span><span>{part.qty}</span><span>{money(part.price * part.qty)}</span></div>)}
            </div>
            <strong>До погодження: {money(approval.totalAmount)} · строк {approval.promisedDate}</strong>
          </div>
        )}
      </section>}
      {canManageRepairFlow && !canShowReadonlyOverview && !canShowManagerView && <section className="process-panel">
        <div className="panel-heading">
          <h2>{isEngineerRole ? 'Ремонт и этапы' : 'Клиентские статусы и действия'}</h2>
          <span>{isEngineerRole ? 'Лише дві дії: взяти і повернути в комірку' : 'Понятные этапы для клиента без изменения логики'}</span>
        </div>
        <div className="simple-action-panel">
          {isEngineerRole && (
            <button type="button" onClick={() => runWithWarnings('Инженер взял заказ в ремонт', () => acceptOrderWork(selectedOrder))} disabled={selectedOrder.engineer !== activeUser.name || Boolean(selectedOrder.engineerAcceptedAt) || acceptWorkCriticalMessages.length > 0}>Взяти в ремонт</button>
          )}
          {canManageClientWork && <button type="button" onClick={() => runWithWarnings('Менеджер отправил согласование', () => sendRepairApproval(selectedOrder, 'Telegram'))} disabled={statusActionCriticalMessages.length > 0}>Відправити на погодження</button>}
          {isEngineerRole && <button type="button" onClick={() => runWithWarnings('Інженер завершив ремонт і повернув у ячейку', () => returnOrderToCellReady(selectedOrder))} disabled={!nextStatuses.includes('Готовий до видачі') || !roleCanSetOrderStatus(activeUser.role, 'Готовий до видачі') || statusActionCriticalMessages.length > 0 || !canCompleteRepair || selectedOrder.locationStatus !== 'У інженера'}>Готово, повернено в ячейку</button>}
          {canManageClientWork && canIssueOrder && <button type="button" onClick={() => runWithWarnings('Менеджер выдал заказ клиенту', () => issueReadyOrder(selectedOrder))} disabled={issueCriticalMessages.length > 0}>Видати замовлення</button>}
          {!isSupervisorRole && <button type="button" onClick={() => runWithWarnings('Сотрудник отметил заказ как не подлежащий ремонту', () => changeOrderStatus(selectedOrder, 'Не підлягає ремонту', 'Причина: ремонт технічно або економічно неможливий.'))} disabled={!roleCanSetOrderStatus(activeUser.role, 'Не підлягає ремонту') || ['Закрито', 'Видано', 'Скасовано'].includes(selectedOrder.status) || statusActionCriticalMessages.length > 0}>Не підлягає ремонту</button>}
        </div>
      </section>}
      {canManageClientWork && !canShowManagerView && ['Готовий до видачі', 'Не підлягає ремонту', 'Очікує оплати', 'Очікує клієнта'].includes(selectedOrder.status) && (
        <section className="panel-subsection">
          <div className="panel-heading">
            <h2>Оплата и выдача</h2>
            <span>{simplePaymentStatus} · осталось {money(totals.debt)}</span>
          </div>
          <div className="details-grid">
            <Info label="Сумма ремонта" value={money(totals.total)} />
            <Info label="Оплачено" value={money(totals.paid)} />
            <Info label="Долг" value={money(totals.debt)} />
            <Info label="Статус оплаты" value={simplePaymentStatus} />
          </div>
          <div className="action-row">
            <button type="button" onClick={() => openNotificationPreview(selectedOrder.status === 'Очікує оплати' ? 'Очікує оплату' : 'Готово до видачі')} disabled={statusActionCriticalMessages.length > 0}>Повідомити клієнта</button>
            {canIssueOrder && <button type="button" onClick={() => runWithWarnings('Менеджер выдал заказ клиенту', () => issueReadyOrder(selectedOrder))} disabled={issueCriticalMessages.length > 0}>Видати замовлення</button>}
          </div>
          {totals.debt > 0 && (
            <div className="action-row payment-actions">
              <input aria-label="Сумма оплаты" type="number" min={1} max={totals.debt || undefined} placeholder={String(totals.debt)} value={paymentAmount || ''} onChange={(event) => setPaymentAmount(Number(event.target.value))} />
              <select aria-label="Способ оплаты" value={paymentMethod} onChange={(event) => setPaymentMethod(event.target.value as PaymentMethod)}>
                <option value="Готівка">Наличные</option>
                <option value="Картка">Карта</option>
                <option value="Безготівка">Безнал</option>
              </select>
              <button type="button" onClick={() => recordOrderPayment(paymentMethod === 'Безготівка' ? 'Безналичная оплата ремонта' : 'Оплата ремонта')} disabled={!canDo('finance.payments')}>Зафиксировать оплату</button>
            </div>
          )}
        </section>
      )}
      {canPreviewOrderDocuments && !canShowEngineerView && !canShowManagerView && (
        <section className="panel-subsection">
          <div className="panel-heading">
            <h2>Документы по заказу</h2>
            <span>Чистые шаблоны для проверки и печати A4</span>
          </div>
          {isAdminRole && totals.debt > 0 && (
            <div className="action-row payment-actions">
              <input aria-label="Сума оплати" type="number" min={1} max={totals.debt || undefined} placeholder={String(totals.debt)} value={paymentAmount || ''} onChange={(event) => setPaymentAmount(Number(event.target.value))} />
              <select aria-label="Тип оплати" value={paymentMethod} onChange={(event) => setPaymentMethod(event.target.value as PaymentMethod)}>
                <option value="Готівка">Готівка</option>
                <option value="Картка">Картка</option>
                <option value="Безготівка">Безготівка</option>
              </select>
              <button type="button" onClick={() => recordOrderPayment(paymentMethod === 'Безготівка' ? 'Безготівкова оплата' : 'Оплата ремонту')} disabled={!canDo('finance.payments')}>Прийняти оплату</button>
              {!isReadyForClient && <button type="button" onClick={() => addOrderPayment(selectedOrder, Math.min(500, totals.debt || 500), 'Готівка', 'Передплата')} disabled={!canDo('finance.payments')}>Прийняти передплату</button>}
            </div>
          )}
          <div className="action-row secondary-actions">
            {canPreviewWorkOrder && (
              <button type="button" onClick={() => openDocumentPreview('Наряд на ремонт')} style={documentButtonStyle('#8b5e34')}>
                <ClipboardList size={16} />
                <span>Печать наряда</span>
              </button>
            )}
            {canPreviewInvoice && !hasDocument('Рахунок на оплату') && (
              <button type="button" onClick={() => createServiceOrderDocument('Рахунок на оплату', selectedOrder)} style={documentButtonStyle('#1f6fb2')}>
                <Banknote size={16} />
                <span>Створити рахунок</span>
              </button>
            )}
            {canPreviewInvoice && hasDocument('Рахунок на оплату') && (
              <button type="button" onClick={() => openDocumentPreview('Рахунок на оплату')} style={documentButtonStyle('#1f6fb2')}>
                <Banknote size={16} />
                <span>Роздрукувати рахунок</span>
              </button>
            )}
            {canPreviewAct && !hasDocument('Акт надання послуг') && (
              <button type="button" onClick={() => createServiceOrderDocument('Акт надання послуг', selectedOrder)} style={documentButtonStyle('#2f855a')}>
                <CheckCircle2 size={16} />
                <span>Створити акт</span>
              </button>
            )}
            {canPreviewAct && hasDocument('Акт надання послуг') && (
              <button type="button" onClick={() => openDocumentPreview('Акт надання послуг')} style={documentButtonStyle('#2f855a')}>
                <CheckCircle2 size={16} />
                <span>Роздрукувати акт</span>
              </button>
            )}
            {canPreviewWarranty && (
              <button type="button" onClick={() => openDocumentPreview('Гарантійний талон')} style={documentButtonStyle('#6b46c1')}>
                <ShieldCheck size={16} />
                <span>Гарантийный талон</span>
              </button>
            )}
            {isAdminRole && canCreateDeliveryNote && !hasDocument('Видаткова накладна') && <button type="button" onClick={() => createServiceOrderDocument('Видаткова накладна', selectedOrder)}>Створити видаткову накладну</button>}
            {isAdminRole && canCreateDeliveryNote && hasDocument('Видаткова накладна') && <button type="button" onClick={() => openDocumentPreview('Видаткова накладна')}>Роздрукувати видаткову накладну</button>}
            {isAdminRole && canFormTechnicalAct && !hasDocument('Акт технічного стану') && <button type="button" onClick={() => printDocument('Акт технічного стану', 'service_order', selectedOrder.id, selectedOrder.client)}>Сформувати акт технічного стану</button>}
            {isAdminRole && canTransferBas && <button type="button" onClick={() => transferOrderToBas(selectedOrder)}>Передати в BAS</button>}
            {!canPreviewAct && !canPreviewInvoice && !canPreviewWarranty && (
              <div className="empty-state">Сейчас доступен наряд. Рахунок появится при сумме заказа, акт на этапе готовности, гарантия после выдачи.</div>
            )}
          </div>
        </section>
      )}
      {canShowDetailedFinance && <section className="panel-subsection">
        <div className="panel-heading"><h2>Финансы заказа</h2><span>видалення і зміна заднім числом заборонені</span></div>
        <div className="table payments-table">
          <div className="table-row table-head"><span>Дата</span><span>Сума</span><span>Тип</span><span>Статус</span><span>Транзакція</span><span>Співробітник</span></div>
          {selectedOrder.payments.map((payment) => (
            <div className="table-row" key={payment.id}>
              <span>{payment.date}</span>
              <span>{money(payment.amount)}</span>
              <span>{payment.method}<small>{payment.type}</small></span>
              <span>{payment.status ?? 'Підтверджено'}{payment.confirmedBy ? <small>{payment.confirmedBy} · {payment.confirmedAt}</small> : null}</span>
              <span>{payment.transactionNo}</span>
              <span>{payment.acceptedBy}</span>
            </div>
          ))}
          {selectedOrder.payments.length === 0 && <div className="empty-state">Оплат ще немає. Оплата додається тільки тут, у картці заказа.</div>}
        </div>
      </section>}
      <section className="panel-subsection">
        <div className="panel-heading">
          <h2>{isEngineerRole ? 'Мой заработок по заказу' : 'Работы и заработок сотрудника'}</h2>
          <span>{money(totalEmployeeEarning)} по этому заказу</span>
        </div>
        <div className="table payroll-table">
          <div className="table-row table-head"><span>Тип работы</span><span>Исполнитель</span><span>Количество</span><span>Ставка</span><span>Сумма работы</span><span>Заработок</span></div>
          {workCompensationRows.map((row) => (
            <div className="table-row" key={row.id}>
              <span>{row.workType}<small>{row.workName}</small></span>
              <span>{row.employee}</span>
              <span>{row.qty}</span>
              <span>{row.rate}</span>
              <span>{money(row.workAmount)}</span>
              <span>{money(row.earning)}</span>
            </div>
          ))}
          {workCompensationRows.length === 0 && <div className="empty-state">Работы ещё не добавлены, начисление пока не считается.</div>}
        </div>
      </section>
      {canShowCommentBlock && <section className="panel-subsection">
        <div className="panel-heading"><h2>{isEngineerRole ? 'Проблема, диагностика и тест' : 'Комментарии по заказу'}</h2><span>коротко по суті</span></div>
        <div className="task-list">
          <Task icon={<ClipboardList />} title="Проблема клієнта" text={selectedOrder.issue} />
          <Task icon={<Wrench />} title="Діагностика" text={selectedOrder.diagnosisResult ?? 'Ще не заповнено'} />
          <Task icon={<CheckCircle2 />} title="Тест / результат" text={selectedOrder.testResult ?? selectedOrder.clientComment ?? 'Коментарів поки немає'} />
        </div>
      </section>}
      {canShowEngineerView && <section className="panel-subsection">
        <div className="panel-heading"><h2>Работы</h2><span>{selectedOrder.works.length} позиций</span></div>
        <div className="task-list">
          {selectedOrder.works.map((work, index) => (
            <Task key={`${work.name}-${index}`} icon={<Wrench />} title={work.name} text={`${workTypeLabel(work)} · ${work.qty ?? 1} шт. · ставка ${workRateLabel(work, payrollRules.find((item) => item.employee === (work.engineer ?? selectedOrder.engineer)))} · заработок ${money(payrollForWork(selectedOrder, work, payrollRules.find((item) => item.employee === (work.engineer ?? selectedOrder.engineer))))}`} />
          ))}
          {selectedOrder.works.length === 0 && <div className="empty-state">Работы ещё не добавлены.</div>}
        </div>
      </section>}
      {canManageClientWork && !canShowReadonlyOverview && <section className="panel-subsection">
        <div className="panel-heading">
          <h2>Уведомления клиенту</h2>
          <span>{clientNotified ? 'Уведомлён' : 'Не уведомлён'}</span>
        </div>
        <div className="details-grid">
          <Info label="Статус" value={clientNotified ? 'Уведомлён' : 'Не уведомлён'} />
          <Info label="Последний шаблон" value={lastOrderNotification ? notificationDisplay(lastOrderNotification.event) : 'Ещё не отправляли'} />
          <Info label="Последняя отправка" value={lastOrderNotification ? lastOrderNotification.createdAt : 'Нет'} />
          <Info label="По выбранному шаблону" value={selectedNotificationStatus ? 'Уведомлён' : 'Не уведомлён'} />
          <Info label="Тип уведомления" value={selectedOrder.lastNotificationType ? notificationDisplay(selectedOrder.lastNotificationType) : 'Не зафіксовано'} />
        </div>
        <div className="action-row">
          <select aria-label="Шаблон уведомления" value={selectedNotificationEvent} onChange={(event) => setSelectedNotificationEvent(event.target.value as NotificationEvent)}>
            {simpleNotificationOptions.map((option) => (
              <option key={option.value} value={option.value}>{option.label}</option>
            ))}
          </select>
          <button type="button" onClick={() => openNotificationPreview()}>Отправить уведомление</button>
          <button type="button" onClick={openExtraApprovalPreview}>Согласовать доп. работы</button>
          <button type="button" onClick={() => lastOrderNotification && setNotificationPreview({ event: lastOrderNotification.event, text: lastOrderNotification.message, channel: lastOrderNotification.channel })} disabled={!lastOrderNotification}>Повторить уведомление</button>
        </div>
        {notificationPreview && (
          <div className="panel-subsection">
            <div className="panel-heading">
              <h2>Текст уведомления</h2>
              <span>{notificationDisplay(notificationPreview.event)}</span>
            </div>
            <div className="task-list">
              <Task icon={<Bell />} title={`${selectedOrder.client} · ${selectedOrder.device}`} text={`Статус: ${statusDisplay(selectedOrder.status)}`} />
            </div>
            <textarea readOnly value={notificationPreview.text} rows={5} aria-label="Текст уведомления" />
            <div className="action-row">
              <button type="button" onClick={() => { void copyNotificationPreview(); }}>Скопировать</button>
              <button type="button" onClick={() => confirmNotificationPreview('Viber', `viber://forward?text=${encodeURIComponent(notificationPreview.text)}`)}>Открыть Viber</button>
              <button type="button" onClick={() => confirmNotificationPreview('Telegram', `https://t.me/share/url?text=${encodeURIComponent(notificationPreview.text)}`)}>Открыть Telegram</button>
              <button type="button" onClick={() => confirmNotificationPreview('SMS', `sms:${selectedOrder.phone}?body=${encodeURIComponent(notificationPreview.text)}`)}>Открыть SMS</button>
              <button type="button" onClick={() => setNotificationPreview(null)}>Закрыть</button>
            </div>
          </div>
        )}
        <div className="task-list">
          {notifications.slice(0, 4).map((message) => (
            <Task key={message.id} icon={<Bell />} title={`${notificationDisplay(message.event)} · ${message.channel} · ${message.status}`} text={`${message.createdAt} · ${message.createdBy ?? 'система'}. ${message.message}${message.error ? ` Помилка: ${message.error}` : ''}`} />
          ))}
          {notifications.length === 0 && <div className="empty-state">Выберите шаблон и отправьте уведомление. CRM зафиксирует это в журнале и отметит заказ как уведомлённый.</div>}
        </div>
      </section>}
      {canShowHistoryBlock && !canShowEngineerView && !canShowManagerView && <section className="panel-subsection">
        <div className="panel-heading"><h2>{isSupervisorRole ? 'История и контроль' : 'История статусов'}</h2><span>{selectedOrder.statusHistory.length} записів</span></div>
        <div className="task-list">
          {selectedOrder.statusHistory.slice(0, 6).map((entry) => (
            <Task key={entry.id} icon={<History />} title={`${entry.oldStatus ?? 'Старт'} -> ${entry.newStatus}`} text={`${entry.changedAt} · ${entry.changedBy}. ${entry.comment}`} />
          ))}
        </div>
      </section>}
      {canShowPartsSection && <div className="panel-heading parts-heading">
        <h2>{isEngineerRole ? 'Запчасти и работы' : 'Запчасти в ремонте'}</h2>
        {!canShowManagerView && !canShowReadonlyOverview && canDo('repairs.close') && <button className="primary-action" type="button" onClick={() => issueReadyOrder(selectedOrder)} disabled={!selectedOrder.locationCode || selectedOrder.locationStatus === 'У інженера'}>Закрити і видати</button>}
      </div>}
      {canShowPartsSection && canAddParts && (
        <div className="inline-part-form">
          <select value={selectedProductId} onChange={(event) => setSelectedProductId(event.target.value)} aria-label="Запчастина">
            {products.map((product) => <option key={product.id} value={product.id}>{product.name} · доступно {available(product)}</option>)}
          </select>
          <input aria-label="Кількість" type="number" min={1} value={qty} onChange={(event) => setQty(Number(event.target.value))} />
          <button className="submit-button" type="button" onClick={addPartToRepair}>Додати запчастину</button>
        </div>
      )}
      {canShowPartsSection && <div className="repair-parts">
        {selectedOrder.parts.length === 0 && <div className="empty-state">Запчастини ще не додані.</div>}
        {selectedOrder.parts.map((part) => {
          const product = products.find((item) => item.id === part.productId);
          return (
            <article className="repair-part-card" key={part.id}>
              <div>
                <strong>{product?.name}</strong>
                <p>{part.qty} шт.{canSeeFinance ? ` · собівартість ${showCost ? money(part.cost * part.qty) : 'приховано'} · продаж ${money(part.price * part.qty)}` : ' · фінанси приховано'}</p>
              </div>
              <span className={statusClass[part.status]}>{part.status}</span>
              {canShowPartActions && <div className="action-row">
                <button type="button" onClick={() => orderPart(selectedOrder, part)} disabled={!canDo('purchases.create') || !['Потрібно', 'До закупівлі'].includes(part.status)}>Замовити запчастину</button>
                <button type="button" onClick={() => reserveArrived(selectedOrder, part)} disabled={!canDo('stock.reserve') || part.status !== 'Прибуло'}>Зарезервувати</button>
                <button type="button" onClick={() => issueToEngineer(selectedOrder, part)} disabled={!canDo('stock.writeoff') || part.status !== 'Зарезервовано'}>Видати інженеру</button>
                <button type="button" onClick={() => markInstalled(selectedOrder, part)} disabled={!canDo('repairs.status') || part.status !== 'Видано інженеру'}>Відмітити встановленою</button>
                <button type="button" onClick={() => returnServicePart(selectedOrder, part, 'На склад')} disabled={!canDo('stock.unreserve') || !['Зарезервовано', 'Видано інженеру', 'Встановлено'].includes(part.status)}>Повернути в сервісі</button>
                <button type="button" onClick={() => returnServicePart(selectedOrder, part, 'Брак')} disabled={!canDo('stock.writeoff') || !['Видано інженеру', 'Встановлено'].includes(part.status)}>Списати в брак</button>
              </div>}
            </article>
          );
        })}
      </div>}
      {isLabelPreviewOpen && (
        <div style={{ position: 'fixed', inset: 0, backgroundColor: 'rgba(15, 23, 42, 0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '24px', zIndex: 1190 }}>
          <div style={{ width: 'min(420px, 100%)', backgroundColor: '#fff', border: '1px solid #cbd5e1', boxShadow: '0 18px 48px rgba(15, 23, 42, 0.28)', padding: '20px' }}>
            <div className="panel-heading">
              <h2>Наклейка замовлення</h2>
              <span>{selectedOrder.id} · {selectedOrder.locationCode ?? 'без комірки'}</span>
            </div>
            <div style={{ border: '2px solid #0f172a', padding: '18px', display: 'grid', gap: '12px' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', gap: '12px', alignItems: 'center' }}>
                <div>
                  <div style={{ fontSize: '30px', fontWeight: 800 }}>{selectedOrder.id}</div>
                  <div style={{ fontSize: '34px', fontWeight: 900, marginTop: '10px', lineHeight: 1 }}>{selectedOrder.locationCode ?? 'БЕЗ ЯЧЕЙКИ'}</div>
                  <div style={{ display: 'inline-block', marginTop: '8px', padding: '6px 10px', color: '#fff', backgroundColor: labelStatusColor, fontWeight: 700 }}>{labelStatusText}</div>
                  <div style={{ display: 'inline-block', marginTop: '8px', padding: '4px 10px', border: '1px solid #0f172a', fontWeight: 700 }}>{labelLocationStateText}</div>
                  {selectedOrder.urgent && <div style={{ display: 'inline-block', marginTop: '8px', padding: '4px 10px', border: '1px solid #b45309', color: '#b45309', fontWeight: 700 }}>ТЕРМІНОВО</div>}
                </div>
                <img src={`https://api.qrserver.com/v1/create-qr-code/?size=180x180&data=${encodeURIComponent(directOrderUrl)}`} alt={`QR ${selectedOrder.id}`} style={{ width: '110px', height: '110px', border: '1px solid #cbd5e1' }} />
              </div>
              <Info label="Дата прийому" value={extractDayKey(selectedOrder.intakeDate)} />
              <Info label="Вік замовлення" value={`${totalAge} дн.`} />
              <Info label="Клієнт" value={selectedOrder.client} />
              <Info label="Пристрій" value={selectedOrder.device} />
            </div>
            <div className="action-row" style={{ justifyContent: 'flex-end', marginTop: '16px' }}>
              <button type="button" onClick={printLabelPreview}>Печать</button>
              <button type="button" onClick={() => setIsLabelPreviewOpen(false)}>Закрыть</button>
            </div>
          </div>
        </div>
      )}
      {isDocumentPreviewOpen && (
        <div style={{ position: 'fixed', inset: 0, backgroundColor: 'rgba(15, 23, 42, 0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '24px', zIndex: 1200 }}>
          <div style={{ width: 'min(1100px, 100%)', maxHeight: '92vh', backgroundColor: '#f4f1ea', border: '1px solid #b5aa98', boxShadow: '0 18px 48px rgba(15, 23, 42, 0.28)', display: 'grid', gridTemplateRows: 'auto auto 1fr auto', overflow: 'hidden' }}>
            <div style={{ padding: '14px 18px 10px', borderBottom: '1px solid #cfc4b2' }}>
              <div className="panel-heading" style={{ marginBottom: 0 }}>
                <h2>{selectedDocumentKind === 'Наряд на ремонт' ? 'Печать наряда' : selectedDocumentKind === 'Рахунок на оплату' ? 'Рахунок' : selectedDocumentKind === 'Акт надання послуг' ? 'Акт выполненных работ' : selectedDocumentKind === 'Видаткова накладна' ? 'Видаткова накладна' : 'Гарантийный талон'}</h2>
                <span>{selectedOrder.id} · {selectedOrder.client} · формат A4</span>
              </div>
            </div>
            <div style={{ padding: '10px 18px', borderBottom: '1px solid #d8cebf', backgroundColor: '#efe9de', fontSize: '13px', color: '#5b4a33' }}>
              Документ открывается сразу в нужном виде. Шаблон A4 и структура печати не изменяются.
            </div>
            <div style={{ padding: '16px', overflow: 'auto', backgroundColor: '#d8d1c2' }}>
              <iframe
                key={`${selectedOrder.id}-${selectedDocumentKind}-${documentPreviewNonce}`}
                ref={documentIframeRef}
                title={`Документ ${selectedDocumentKind}`}
                srcDoc={orderDocumentHtml}
                style={{ width: '100%', minHeight: '72vh', border: '1px solid #9f9789', backgroundColor: '#fff' }}
              />
            </div>
            <div className="action-row" style={{ padding: '12px 18px 16px', borderTop: '1px solid #d8cebf', backgroundColor: '#efe9de', justifyContent: 'flex-end' }}>
              <button type="button" onClick={printPreviewDocument}>Печать</button>
              <button type="button" onClick={() => setIsDocumentPreviewOpen(false)}>Закрыть</button>
            </div>
          </div>
        </div>
      )}
    </section>
  );
}

function SalesPage(props: {
  sales: Sale[];
  selectedSale: Sale;
  products: Product[];
  selectedSaleId: string;
  setSelectedSaleId: (id: string) => void;
  selectedSaleProductId: string;
  setSelectedSaleProductId: (id: string) => void;
  saleQty: number;
  setSaleQty: (qty: number) => void;
  addSaleItem: () => void;
  reserveSale: (sale: Sale) => void;
  acceptSalePayment: (sale: Sale, amount: number, method: PaymentMethod, comment: string) => void;
  issueSale: (sale: Sale) => void;
  returnSale: (sale: Sale) => void;
  cancelSale: (sale: Sale) => void;
  printDocument: (kind: DocumentKind, entityType: PrintDocument['entityType'], entityId: string, clientOrSupplier: string) => void;
  canDo: (permission: Permission) => boolean;
  showCost: boolean;
  showProfit: boolean;
}) {
  const totals = saleTotals(props.selectedSale);
  const [showSaleMore, setShowSaleMore] = useState(false);
  const [saleScanCode, setSaleScanCode] = useState('');
  const scannedSaleProduct = findProductByLookup(props.products, saleScanCode.trim());
  return (
    <div className="page-grid">
      <PageTitle eyebrow="Продажі" title="Продаж → резерв → оплата → видача → списання" text="Товар списується тільки після підтвердженої видачі. До цього він або в чернетці, або в резерві." />
      <section className="content-split orders-layout">
        <div className="panel">
          <div className="panel-heading"><h2>Продажі</h2><span>{props.sales.length} документів</span></div>
          <div className="order-list">
            {props.sales.map((sale) => {
              const saleInfo = saleTotals(sale);
              return (
                <button type="button" className={props.selectedSaleId === sale.id ? 'order-card order-card-active' : 'order-card'} key={sale.id} onClick={() => props.setSelectedSaleId(sale.id)}>
                  <div className="order-main">
                    <div>
                      <div className="order-title"><strong>{sale.id}</strong><span className={statusClass[sale.status]}>{sale.status}</span></div>
                      <h3>{sale.client}</h3>
                      <p>{sale.manager} · {sale.date}</p>
                    </div>
                    <div className="order-price"><strong>{money(saleInfo.total)}</strong><span>борг {money(saleInfo.debt)}</span></div>
                  </div>
                </button>
              );
            })}
          </div>
        </div>
        <div className="panel order-form">
          <div className="panel-heading"><h2>Додати товар</h2></div>
          <label>
            Сканер штрих-коду
            <input
              value={saleScanCode}
              placeholder="Скануйте barcode або введіть артикул"
              onChange={(event) => {
                const code = event.target.value;
                setSaleScanCode(code);
                const product = findProductByLookup(props.products, code.trim());
                if (product) props.setSelectedSaleProductId(product.id);
              }}
            />
            <small>{scannedSaleProduct ? `Знайдено: ${scannedSaleProduct.name}` : 'Сканер працює як клавіатура: наведіть курсор у поле і скануйте.'}</small>
          </label>
          <label>
            Товар
            <select value={props.selectedSaleProductId} onChange={(event) => props.setSelectedSaleProductId(event.target.value)}>
              {props.products.map((product) => <option key={product.id} value={product.id}>{product.name} · доступно {available(product)}</option>)}
            </select>
          </label>
          <label>
            Кількість
            <input type="number" min={1} value={props.saleQty} onChange={(event) => props.setSaleQty(Number(event.target.value))} />
          </label>
          <button className="submit-button" type="button" onClick={props.addSaleItem} disabled={!props.canDo('sales.create')}>Додати товар</button>
        </div>
      </section>

      <section className="panel">
        <div className="panel-heading">
          <h2>{props.selectedSale.id} · {props.selectedSale.client}</h2>
          <span className={statusClass[props.selectedSale.status]}>{props.selectedSale.status}</span>
        </div>
        <div className="details-grid">
          <Info label="Менеджер" value={props.selectedSale.manager} />
          <Info label="Дата" value={props.selectedSale.date} />
          <Info label="Пов'язаний ремонт" value={props.selectedSale.serviceOrderId ?? "Не прив'язано"} />
        </div>
        <section className="stats-grid">
          <Metric icon={<Banknote />} label="Сума продажу" value={money(totals.total)} hint={`знижка ${money(totals.discount)}`} />
          <Metric icon={<CheckCircle2 />} label="Оплачено" value={money(totals.paid)} hint={`борг ${money(totals.debt)}`} />
          <Metric icon={<Archive />} label="Собівартість" value={props.showCost ? money(totals.cost) : 'Приховано'} hint="після видачі" />
          <Metric icon={<PackageCheck />} label="Прибуток" value={props.showProfit ? money(totals.profit) : 'Приховано'} hint="з урахуванням повернень" />
        </section>
        <div className="repair-parts">
          {props.selectedSale.items.length === 0 && <div className="empty-state">Додайте товар у продаж.</div>}
          {props.selectedSale.items.map((item) => {
            const product = props.products.find((entry) => entry.id === item.productId);
            return (
              <article className="repair-part-card" key={item.id}>
                <div>
                  <strong>{product?.name}</strong>
                  <p>{item.qty} шт. · ціна {money(item.price)} · собівартість {props.showCost ? money(item.cost) : 'приховано'} · знижка {money(item.discount)}</p>
                </div>
                <span className={statusClass[item.status === 'Чернетка' ? 'Чернетка' : item.status === 'Скасовано' ? 'Скасовано' : item.status === 'Повернення' ? 'Повернення' : item.status === 'Видано' ? 'Видано' : 'Зарезервовано']}>{item.status}</span>
              </article>
            );
          })}
        </div>
        <div className="action-row payment-actions">
          <button type="button" onClick={() => props.reserveSale(props.selectedSale)} disabled={!props.canDo('stock.reserve')}>Зарезервувати</button>
          <button type="button" onClick={() => props.acceptSalePayment(props.selectedSale, totals.debt, 'Готівка', 'Оплата готівкою')} disabled={!props.canDo('finance.payments')}>Прийняти оплату</button>
          <button type="button" onClick={() => props.issueSale(props.selectedSale)} disabled={!props.canDo('sales.post')}>Видати товар</button>
          <button type="button" onClick={() => props.printDocument('Товарний чек', 'sale', props.selectedSale.id, props.selectedSale.client)}>Сформувати чек</button>
          <button className="ghost-toggle compact-toggle" type="button" onClick={() => setShowSaleMore((value) => !value)}>{showSaleMore ? 'Сховати додаткові дії' : 'Показати ще'}</button>
        </div>
        {showSaleMore && <div className="action-row secondary-actions">
          <button type="button" onClick={() => props.acceptSalePayment(props.selectedSale, totals.debt, 'Картка', 'Оплата карткою')} disabled={!props.canDo('finance.payments')}>Провести оплату карткою</button>
          <button type="button" onClick={() => props.acceptSalePayment(props.selectedSale, Math.max(Math.round(totals.debt / 2), 1), 'Готівка', 'Часткова оплата')} disabled={!props.canDo('finance.payments')}>Відмітити часткову оплату</button>
          <button type="button" onClick={() => props.printDocument('Накладна продажу', 'sale', props.selectedSale.id, props.selectedSale.client)}>Сформувати накладну</button>
          {props.selectedSale.items.length > 0 && <button type="button" onClick={() => props.printDocument('Видаткова накладна', 'sale', props.selectedSale.id, props.selectedSale.client)}>Створити видаткову</button>}
          <button type="button" onClick={() => props.returnSale(props.selectedSale)} disabled={!props.canDo('sales.return') || !props.canDo('finance.refund')}>Зробити повернення</button>
          <button type="button" onClick={() => props.cancelSale(props.selectedSale)} disabled={!props.canDo('sales.cancel')}>Скасувати продаж</button>
        </div>}
        <div className="panel-heading parts-heading"><h2>Платежі</h2><span>{props.selectedSale.payments.length} операцій</span></div>
        <div className="task-list">
          {props.selectedSale.payments.length === 0 && <div className="empty-state">Оплат ще немає.</div>}
          {props.selectedSale.payments.map((payment) => (
            <Task key={payment.id} icon={<Banknote />} title={`${payment.type} · ${payment.method} · ${money(payment.amount)}`} text={`Транзакція: ${payment.transactionNo}. Прийняв: ${payment.acceptedBy}. ${payment.comment}`} />
          ))}
        </div>
      </section>
    </div>
  );
}

function PartsPage({ products, requirements, receipts, movements, productName, showCost, receiveStockIntake, canDo, onImportProducts }: { products: Product[]; requirements: PartRequirement[]; receipts: GoodsReceipt[]; movements: StockMovement[]; productName: (id: string) => string; showCost: boolean; receiveStockIntake: (input: StockIntakeInput) => void; canDo: (permission: Permission) => boolean; onImportProducts: (file: File) => Promise<void> }) {
  const canReceiveStock = canDo('stock.receive');
  const [stockScanCode, setStockScanCode] = useState('');
  const [stockSearch, setStockSearch] = useState('');
  const [expandedProductId, setExpandedProductId] = useState('');
  const [intakeMode, setIntakeMode] = useState<'existing' | 'new'>('existing');
  const [intakeProductId, setIntakeProductId] = useState(products[0]?.id ?? '');
  const [intakeName, setIntakeName] = useState('');
  const [intakeSku, setIntakeSku] = useState('');
  const [intakeBrand, setIntakeBrand] = useState('');
  const [intakeModel, setIntakeModel] = useState('');
  const [intakeBarcode, setIntakeBarcode] = useState('');
  const [intakeQty, setIntakeQty] = useState('1');
  const [intakePrice, setIntakePrice] = useState('');
  const [intakeSupplier, setIntakeSupplier] = useState('');
  const [intakeCategory, setIntakeCategory] = useState('');
  const [intakeUnit, setIntakeUnit] = useState('шт');
  const [intakeMinStock, setIntakeMinStock] = useState('1');
  const [intakeLocation, setIntakeLocation] = useState('');
  const [intakeDocumentNo, setIntakeDocumentNo] = useState('');
  const [duplicatePayload, setDuplicatePayload] = useState<StockIntakeInput | null>(null);
  const [duplicateCandidates, setDuplicateCandidates] = useState<Product[]>([]);
  const importProductsInputRef = useRef<HTMLInputElement | null>(null);
  const scannedProduct = findProductByLookup(products, stockScanCode.trim());
  const sortedProducts = [...products].sort((a, b) => {
    const availableDelta = available(a) - available(b);
    if (availableDelta !== 0) return availableDelta;
    return a.name.localeCompare(b.name, 'uk');
  });
  const filteredProducts = sortedProducts.filter((product) => {
    const needle = stockSearch.trim().toLowerCase();
    if (!needle) return true;
    return [
      product.sku,
      product.name,
      product.barcode,
      ...product.extraBarcodes,
    ].join(' ').toLowerCase().includes(needle);
  });

  function getStockTone(product: Product) {
    const free = available(product);
    if (free < Math.max(product.min, 1)) return { tone: 'danger' as const, label: 'Нижче мінімуму' };
    if (free <= Math.max(product.min + 2, 3)) return { tone: 'warning' as const, label: 'Мало' };
    return { tone: 'normal' as const, label: 'Норма' };
  }

  function submitIntake() {
    if (intakeMode === 'existing' && !intakeProductId) return;
    if (intakeMode === 'new' && !intakeSku.trim()) {
      window.alert('Для нової картки товару потрібен SKU.');
      return;
    }
    const selectedProduct = products.find((product) => product.id === intakeProductId);
    const payload: StockIntakeInput = {
      productId: intakeMode === 'existing' ? intakeProductId : undefined,
      name: intakeMode === 'existing' ? (selectedProduct?.name ?? '') : intakeName,
      qty: Number(intakeQty),
      purchasePrice: Number(intakePrice),
      supplier: intakeSupplier,
      category: intakeCategory,
      sku: intakeMode === 'new' ? intakeSku : selectedProduct?.sku,
      brand: intakeMode === 'new' ? intakeBrand : selectedProduct?.brand,
      model: intakeMode === 'new' ? intakeModel : selectedProduct?.model,
      barcode: intakeBarcode,
      unit: intakeMode === 'new' ? intakeUnit : selectedProduct?.unit,
      minStock: Number(intakeMinStock),
      storageLocation: intakeMode === 'new' ? intakeLocation : selectedProduct?.storageLocation,
      purchaseDocumentNo: intakeDocumentNo,
    };
    if (intakeMode === 'new') {
      const possibleDuplicates = findSimilarProducts(products, {
        name: intakeName,
        sku: intakeSku,
        barcode: intakeBarcode,
        brand: intakeBrand,
        model: intakeModel,
      });
      if (possibleDuplicates.length > 0) {
        setDuplicatePayload(payload);
        setDuplicateCandidates(possibleDuplicates);
        return;
      }
    }
    receiveStockIntake(payload);
    setIntakeQty('1');
    setIntakePrice('');
    setIntakeSupplier('');
    setIntakeBarcode('');
    setIntakeDocumentNo('');
    setDuplicatePayload(null);
    setDuplicateCandidates([]);
    if (intakeMode === 'new') {
      setIntakeName('');
      setIntakeSku('');
      setIntakeBrand('');
      setIntakeModel('');
      setIntakeCategory('');
      setIntakeLocation('');
      setIntakeMode('existing');
    }
  }

  function applyDuplicateDecision(kind: 'existing' | 'append-barcode' | 'new', targetProductId?: string) {
    if (!duplicatePayload) return;
    receiveStockIntake({
      ...duplicatePayload,
      productId: kind === 'existing' ? targetProductId : undefined,
      appendBarcodeToProductId: kind === 'append-barcode' ? targetProductId : undefined,
      forceCreateNew: kind === 'new',
    });
    setDuplicatePayload(null);
    setDuplicateCandidates([]);
    setIntakeQty('1');
    setIntakePrice('');
    setIntakeSupplier('');
    setIntakeBarcode('');
    setIntakeDocumentNo('');
    setIntakeName('');
    setIntakeSku('');
    setIntakeBrand('');
    setIntakeModel('');
    setIntakeCategory('');
    setIntakeLocation('');
    setIntakeMode('existing');
  }

  async function handleProductsFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      await onImportProducts(file);
    } catch (error) {
      window.alert(error instanceof Error ? error.message : 'Не вдалося імпортувати склад.');
    } finally {
      event.target.value = '';
    }
  }

  return (
    <div className="page-grid">
      <PageTitle eyebrow="Склад" title="Прихід, залишки і рух товару" text="Склад живе подіями: товар прийшов, зарезервувався, поїхав у ремонт або списався. Ручного редагування залишків немає." />
      {canReceiveStock && (
      <section className="panel">
        <div className="panel-heading"><h2>Прийняти товар</h2><span>швидкий прихід без окремої таблиці</span></div>
        <input ref={importProductsInputRef} type="file" accept=".xlsx,.csv" onChange={handleProductsFileChange} className="backup-hidden-input" />
        <div className="quick-actions warehouse-intake-actions">
          <button type="button" onClick={() => importProductsInputRef.current?.click()}>Імпорт складу</button>
          <button type="button" className={intakeMode === 'existing' ? 'primary' : ''} onClick={() => setIntakeMode('existing')}>Існуюча позиція</button>
          <button type="button" className={intakeMode === 'new' ? 'primary' : ''} onClick={() => setIntakeMode('new')}>Нова позиція</button>
        </div>
        <div className="table">
          {intakeMode === 'existing' ? (
            <label>
              Товар
              <select value={intakeProductId} onChange={(event) => setIntakeProductId(event.target.value)}>
                {products.map((product) => <option key={product.id} value={product.id}>{product.name} · {product.sku}</option>)}
              </select>
            </label>
          ) : (
            <>
              <label>
                Назва товару
                <input value={intakeName} placeholder="Людська назва, не випадковий текст" onChange={(event) => setIntakeName(event.target.value)} />
              </label>
              <label>
                Категорія
                <input value={intakeCategory} placeholder="Наприклад: Клавіатури" onChange={(event) => setIntakeCategory(event.target.value)} />
              </label>
              <label>
                Бренд
                <input value={intakeBrand} placeholder="HP / Logitech / GP" onChange={(event) => setIntakeBrand(event.target.value)} />
              </label>
              <label>
                Модель / характеристика
                <input value={intakeModel} placeholder="K120 USB / X200 Black" onChange={(event) => setIntakeModel(event.target.value)} />
              </label>
              <label>
                Артикул / SKU
                <input value={intakeSku} placeholder="Внутрішній код" onChange={(event) => setIntakeSku(event.target.value)} />
              </label>
              <label>
                Штрих-код
                <input value={intakeBarcode} placeholder="Основний або новий штрих-код" onChange={(event) => setIntakeBarcode(event.target.value)} />
              </label>
              <label>
                Одиниця
                <input value={intakeUnit} placeholder="шт / м / компл" onChange={(event) => setIntakeUnit(event.target.value)} />
              </label>
              <label>
                Мінімальний залишок
                <input type="number" min={0} value={intakeMinStock} onChange={(event) => setIntakeMinStock(event.target.value)} />
              </label>
              <label>
                Комірка / місце
                <input value={intakeLocation} placeholder="A-01-02" onChange={(event) => setIntakeLocation(event.target.value)} />
              </label>
            </>
          )}
          <label>
            Кількість
            <input type="number" min={1} value={intakeQty} onChange={(event) => setIntakeQty(event.target.value)} />
          </label>
          <label>
            Закупівельна ціна
            <input type="number" min={1} value={intakePrice} placeholder="0" onChange={(event) => setIntakePrice(event.target.value)} />
          </label>
          <label>
            Постачальник
            <input value={intakeSupplier} placeholder="Опціонально" onChange={(event) => setIntakeSupplier(event.target.value)} />
          </label>
          <label>
            Документ закупівлі
            <input value={intakeDocumentNo} placeholder="Накладна / рахунок" onChange={(event) => setIntakeDocumentNo(event.target.value)} />
          </label>
        </div>
        <div className="action-row">
          <button type="button" className="submit-button" onClick={submitIntake}>Прийняти товар</button>
          <span className="warehouse-intake-note">Кожен прихід створює окрему партію. Списання далі піде FIFO.</span>
        </div>
        {duplicateCandidates.length > 0 && duplicatePayload && (
          <div className="panel-subsection">
            <div className="panel-heading"><h2>Можливо, такий товар уже є</h2><span>виберіть правильну дію</span></div>
            <div className="task-list">
              {duplicateCandidates.map((product) => (
                <Task
                  key={product.id}
                  icon={<Archive />}
                  title={`${product.name} · ${product.sku}`}
                  text={`Бренд ${product.brand || '—'} · модель ${product.model || '—'} · штрих-коди ${productBarcodeList(product).join(', ') || 'немає'} · залишок ${available(product)} ${product.unit}`}
                />
              ))}
            </div>
            <div className="action-row">
              {duplicateCandidates.map((product) => (
                <button key={`${product.id}-existing`} type="button" onClick={() => applyDuplicateDecision('existing', product.id)}>Використати {product.sku}</button>
              ))}
              {duplicatePayload.barcode && duplicateCandidates.map((product) => (
                <button key={`${product.id}-append`} type="button" onClick={() => applyDuplicateDecision('append-barcode', product.id)}>Додати штрих-код до {product.sku}</button>
              ))}
              <button type="button" onClick={() => applyDuplicateDecision('new')}>Створити новий товар</button>
            </div>
          </div>
        )}
      </section>
      )}
      <section className="panel">
        <div className="panel-heading"><h2>Сканер штрих-коду</h2><span>barcode / артикул</span></div>
        <label>
          Скануйте товар
          <input value={stockScanCode} placeholder="Наприклад 4820001115008 або SSD-KIN-500" onChange={(event) => setStockScanCode(event.target.value)} />
          <small>USB-сканер працює як клавіатура. Поставте курсор у поле і скануйте етикетку.</small>
        </label>
        {scannedProduct && (
          <div className="scan-result">
            <strong>{scannedProduct.name}</strong>
            <span>{scannedProduct.sku} · штрих-коди {productBarcodeList(scannedProduct).join(', ')} · доступно {available(scannedProduct)} {scannedProduct.unit} · партій {scannedProduct.batches.length}</span>
            {canReceiveStock && (
              <div className="action-row warehouse-scan-actions">
                <button type="button" onClick={() => { setIntakeMode('existing'); setIntakeProductId(scannedProduct.id); window.scrollTo({ top: 0, behavior: 'smooth' }); }}>Оприбуткувати в цю картку</button>
              </div>
            )}
          </div>
        )}
        {!scannedProduct && stockScanCode.trim() && (
          <div className="empty-state">Товар не знайдено. Можна створити нову картку через прихід або прив'язати штрих-код до існуючої.</div>
        )}
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Залишки на складі</h2><span>{filteredProducts.length} із {products.length}</span></div>
        <div className="stock-table-search">
          <input
            value={stockSearch}
            onChange={(event) => setStockSearch(event.target.value)}
            placeholder="Пошук: SKU / назва / штрих-код"
          />
        </div>
        <div className="table stock-table stock-table-products">
          <div className="table-row table-head">
            <span>SKU</span>
            <span>Назва</span>
            <span>Залишок</span>
            <span>Мін.</span>
            <span>Комірка</span>
            <span>Сигнал</span>
          </div>
          {filteredProducts.map((product) => {
            const tone = getStockTone(product);
            const isExpanded = expandedProductId === product.id;
            const productReceipts = receipts.filter((receipt) => receipt.productId === product.id).slice(0, 5);
            const productMovements = movements.filter((movement) => movement.productId === product.id).slice(0, 8);
            return (
              <React.Fragment key={product.id}>
                <button
                  type="button"
                  className={`table-row stock-product-row stock-product-row-${tone.tone}${isExpanded ? ' is-open' : ''}`}
                  onClick={() => setExpandedProductId((current) => current === product.id ? '' : product.id)}
                >
                  <span>{product.sku}</span>
                  <span>{product.name}</span>
                  <span>{available(product)} {product.unit}</span>
                  <span>{product.min}</span>
                  <span>{product.storageLocation || '—'}</span>
                  <span className={tone.tone === 'danger' ? 'tag tag-red' : tone.tone === 'warning' ? 'tag tag-yellow' : 'tag'}>{tone.label}</span>
                </button>
                {isExpanded && (
                  <div className="stock-product-detail">
                    <div className="stock-product-detail-block">
                      <strong>{product.name}</strong>
                      <div className="stock-product-inline-meta">
                        <span>SKU {product.sku}</span>
                        <span>{product.brand || '—'} {product.model || ''}</span>
                        <span>Штрих-коди: {productBarcodeList(product).join(', ') || '—'}</span>
                      </div>
                    </div>
                    <div className="stock-product-detail-grid">
                      <div className="stock-product-detail-block">
                        <strong>Партії FIFO</strong>
                        <div className="table stock-detail-table">
                          <div className="table-row table-head">
                            <span>Партія</span>
                            <span>Дата</span>
                            <span>Залишок</span>
                            <span>Ціна</span>
                            <span>Постачальник</span>
                          </div>
                          {product.batches.length > 0 ? product.batches.map((batch) => (
                            <div key={batch.id} className="table-row">
                              <span>{batch.id}</span>
                              <span>{batch.purchaseDate}</span>
                              <span>{batch.qtyAvailable}/{batch.qtyTotal}</span>
                              <span>{showCost ? money(batch.purchasePrice) : 'Приховано'}</span>
                              <span>{batch.supplier || batch.source || '—'}</span>
                            </div>
                          )) : <div className="empty-state">Партій немає.</div>}
                        </div>
                      </div>
                      <div className="stock-product-detail-block">
                        <strong>Закупки</strong>
                        <div className="table stock-detail-table">
                          <div className="table-row table-head">
                            <span>Дата</span>
                            <span>Кількість</span>
                            <span>Ціна</span>
                            <span>Постачальник</span>
                          </div>
                          {productReceipts.length > 0 ? productReceipts.map((receipt) => (
                            <div key={receipt.id} className="table-row">
                              <span>{receipt.date}</span>
                              <span>{receipt.qty}</span>
                              <span>{showCost ? money(receipt.price) : 'Приховано'}</span>
                              <span>{receipt.supplier || '—'}</span>
                            </div>
                          )) : <div className="empty-state">Закупок ще немає.</div>}
                        </div>
                      </div>
                    </div>
                    <div className="stock-product-detail-block">
                      <strong>Рух товару</strong>
                      <div className="table stock-detail-table movement-table">
                        <div className="table-row table-head">
                          <span>Дата</span>
                          <span>Тип</span>
                          <span>Кількість</span>
                          <span>Партія</span>
                          <span>Підстава</span>
                          <span>Коментар</span>
                        </div>
                        {productMovements.length > 0 ? productMovements.map((movement) => (
                          <div key={movement.id} className="table-row">
                            <span>{movement.date}</span>
                            <span>{movement.type}</span>
                            <span>{movement.qty}</span>
                            <span>{movement.batchRefs || '—'}</span>
                            <span>{movement.basis || movement.orderId || movement.purchaseId || '—'}</span>
                            <span>{movement.comment}</span>
                          </div>
                        )) : <div className="empty-state">Руху ще немає.</div>}
                      </div>
                    </div>
                    {canReceiveStock && (
                      <div className="action-row">
                        <button type="button" onClick={() => { setIntakeMode('existing'); setIntakeProductId(product.id); setIntakePrice(String(product.cost || '')); setIntakeBarcode(''); window.scrollTo({ top: 0, behavior: 'smooth' }); }}>Додати прихід</button>
                      </div>
                    )}
                  </div>
                )}
              </React.Fragment>
            );
          })}
        </div>
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Потреби із ремонтів</h2><span>{requirements.length} позицій</span></div>
        <div className="task-list">
          {requirements.map((req) => <Task key={req.id} icon={<PackagePlus />} title={`${productName(req.productId)} · ${req.qty} шт.`} text={`${req.orderId} · статус: ${req.status}`} />)}
        </div>
      </section>
    </div>
  );
}

function PurchasesPage({
  purchases,
  requirements,
  products,
  productName,
  receivePurchase,
  createManualPurchaseRequest,
  createPurchaseFromRequirement,
  updatePurchaseProcurementMeta,
  printDocument,
  canDo,
}: {
  purchases: PurchaseOrder[];
  requirements: PartRequirement[];
  products: Product[];
  productName: (id: string) => string;
  receivePurchase: (id: string) => void;
  createManualPurchaseRequest: (input: { productId: string; qty: number; reason: PurchaseReason; priority: PurchasePriority; comment: string }) => void;
  createPurchaseFromRequirement: (requirementId: string) => void;
  updatePurchaseProcurementMeta: (purchaseId: string, patch: Partial<PurchaseOrder>) => void;
  printDocument: (kind: DocumentKind, entityType: PrintDocument['entityType'], entityId: string, clientOrSupplier: string) => void;
  canDo: (permission: Permission) => boolean;
}) {
  const canCreatePurchase = canDo('purchases.create');
  const canEditPurchase = canDo('purchases.edit');
  const canChangePurchaseStatus = canDo('purchases.status');
  const canReceiveStock = canDo('stock.receive');
  const [expandedProductId, setExpandedProductId] = useState('');
  const openRequirements = requirements.filter((item) => ['Потрібно', 'До закупівлі', 'Замовлено', 'В дорозі'].includes(item.status));
  const shortageRows = products
    .filter((product) => available(product) < Math.max(product.min, 1))
    .map((product) => {
      const requirement = openRequirements.find((item) => item.productId === product.id);
      const purchase = purchases.find((item) =>
        ['Нова', 'В роботі', 'Замовлено', 'В дорозі'].includes(item.status)
        && item.items.some((line) => line.productId === product.id),
      );
      return {
        product,
        requirement,
        purchase,
        needQty: Math.max(product.min - available(product), 1),
      };
    })
    .sort((a, b) => (available(a.product) - a.product.min) - (available(b.product) - b.product.min) || a.product.name.localeCompare(b.product.name, 'uk'));

  return (
    <div className="page-grid">
      <PageTitle eyebrow="Закупівлі" title="Закупка" text="Швидкий відгук на нестачу товару." />
      <section className="panel">
        <div className="panel-heading"><h2>Потреба в закупці</h2><span>{shortageRows.length}</span></div>
        <div className="table stock-table purchase-table-products">
          <div className="table-row table-head">
            <span>SKU</span>
            <span>Назва</span>
            <span>Залишок</span>
            <span>Мін.</span>
            <span>Докупити</span>
            <span>Дія</span>
          </div>
          {shortageRows.map(({ product, requirement, purchase, needQty }) => {
            const isExpanded = expandedProductId === product.id;
            const purchaseItem = purchase?.items.find((item) => item.productId === product.id);
            const purchaseStatusLabel = purchase?.status === 'В роботі'
              ? 'Нова'
              : purchase?.status === 'Замовлено'
                ? 'Заказано'
                : purchase?.status === 'В дорозі'
                  ? 'В дорозі'
                  : purchase?.status === 'Отримано' || purchase?.status === 'Прибуло'
                    ? 'Отримано'
                    : 'Нова';
            return (
              <React.Fragment key={product.id}>
                <button
                  type="button"
                  className={`table-row stock-product-row stock-product-row-${available(product) < Math.max(product.min, 1) ? 'danger' : 'warning'}${isExpanded ? ' is-open' : ''}`}
                  onClick={() => setExpandedProductId((current) => current === product.id ? '' : product.id)}
                >
                  <span>{product.sku}</span>
                  <span>{product.name}</span>
                  <span>{available(product)} {product.unit}</span>
                  <span>{product.min}</span>
                  <span>{needQty} {product.unit}</span>
                  <span>
                    {purchase
                      ? <span className={purchase.status === 'В дорозі' ? 'tag tag-blue' : purchase.status === 'Отримано' || purchase.status === 'Прибуло' ? 'tag tag-teal' : 'tag tag-yellow'}>{purchaseStatusLabel}</span>
                      : <span className="tag tag-red">Створити закупівлю</span>}
                  </span>
                </button>
                {isExpanded && (
                  <div className="stock-product-detail">
                    {!purchase && canCreatePurchase && (
                      <div className="action-row">
                        <button
                          type="button"
                          className="submit-button"
                          onClick={() => {
                            if (requirement) {
                              createPurchaseFromRequirement(requirement.id);
                            } else {
                              createManualPurchaseRequest({ productId: product.id, qty: needQty, reason: 'Мінімум', priority: 'Високий', comment: 'Автоматична потреба зі складу.' });
                            }
                          }}
                        >
                          Створити закупівлю
                        </button>
                      </div>
                    )}
                    {purchase && (
                      <>
                        <div className="table purchase-detail-form">
                          <label>
                            Постачальник
                            <input value={purchase.supplier} onChange={(event) => updatePurchaseProcurementMeta(purchase.id, { supplier: event.target.value })} readOnly={!canEditPurchase} />
                          </label>
                          <label>
                            Ціна
                            <input
                              type="number"
                              min={1}
                              value={purchaseItem?.price ?? ''}
                              readOnly={!canEditPurchase}
                              onChange={(event) => updatePurchaseProcurementMeta(purchase.id, {
                                items: purchase.items.map((item) => item.productId === product.id ? { ...item, price: Number(event.target.value) || 0 } : item),
                              })}
                            />
                          </label>
                          <label>
                            Кількість
                            <input
                              type="number"
                              min={1}
                              value={purchaseItem?.qty ?? needQty}
                              readOnly={!canEditPurchase}
                              onChange={(event) => updatePurchaseProcurementMeta(purchase.id, {
                                items: purchase.items.map((item) => item.productId === product.id ? { ...item, qty: Number(event.target.value) || 0 } : item),
                              })}
                            />
                          </label>
                          <label>
                            Посилання
                            <input value={purchase.sourceLink ?? ''} onChange={(event) => updatePurchaseProcurementMeta(purchase.id, { sourceLink: event.target.value })} readOnly={!canEditPurchase} />
                          </label>
                          <label>
                            Коментар
                            <input value={purchase.comment ?? ''} onChange={(event) => updatePurchaseProcurementMeta(purchase.id, { comment: event.target.value })} readOnly={!canEditPurchase} />
                          </label>
                        </div>
                        <div className="action-row">
                          {canChangePurchaseStatus && <button type="button" onClick={() => updatePurchaseProcurementMeta(purchase.id, { status: 'Нова' })}>Нова</button>}
                          {canChangePurchaseStatus && <button type="button" onClick={() => updatePurchaseProcurementMeta(purchase.id, { status: 'Замовлено' })}>Заказано</button>}
                          {canChangePurchaseStatus && <button type="button" onClick={() => updatePurchaseProcurementMeta(purchase.id, { status: 'В дорозі' })}>В дорозі</button>}
                          {canReceiveStock && <button type="button" className="submit-button" onClick={() => receivePurchase(purchase.id)} disabled={!(purchaseItem?.price && purchaseItem.price > 0)}>Отримано</button>}
                          {(canReceiveStock || canEditPurchase || canChangePurchaseStatus) && <button type="button" onClick={() => printDocument('Прихідна накладна', 'purchase_order', purchase.id, purchase.supplier || 'Постачальник')}>Прихідна</button>}
                        </div>
                      </>
                    )}
                  </div>
                )}
              </React.Fragment>
            );
          })}
          {shortageRows.length === 0 && <div className="empty-state">Усі позиції в нормі.</div>}
        </div>
      </section>
    </div>
  );
}

function CashPage({
  orders,
  sales,
  cashShift,
  activeUser,
  confirmBankPayment,
  openCashShift,
  closeCashShift,
  cancelPayment,
  refundPayment,
}: {
  orders: ServiceOrder[];
  sales: Sale[];
  cashShift: CashShift;
  activeUser: User;
  confirmBankPayment: (paymentId: string) => void;
  openCashShift: (openingCashInput: string) => boolean;
  closeCashShift: (actualCashInput: string, comment: string) => boolean;
  cancelPayment: (paymentId: string, reason: string, comment?: string) => void;
  refundPayment: (paymentId: string, amountInput: string, reason: string, method: PaymentMethod) => void;
}) {
  const [openingCash, setOpeningCash] = useState(String(cashShift.openingCash || '0'));
  const [actualCash, setActualCash] = useState('');
  const [shiftComment, setShiftComment] = useState('');
  const [paymentActionDraft, setPaymentActionDraft] = useState<null | {
    paymentId: string;
    action: 'cancel' | 'refund';
    amount: string;
    reason: string;
    method: PaymentMethod;
  }>(null);
  const salePayments = sales.flatMap((sale) => sale.payments.map((payment) => ({ ...payment, client: sale.client, document: sale.id, kind: 'sale' as const })));
  const orderPayments = orders.flatMap((order) => order.payments.map((payment) => ({ ...payment, client: order.client, document: order.id, kind: 'order' as const })));
  const payments = [...salePayments, ...orderPayments].sort((a, b) => b.date.localeCompare(a.date));
  const expectedCash = expectedCashAmount(cashShift);
  const unconfirmedPayments = payments.filter((payment) => paymentNeedsConfirmation(payment));
  const closeDifferencePreview = actualCash.trim() ? Number(actualCash) - expectedCash : 0;

  useEffect(() => {
    setOpeningCash(String(cashShift.openingCash || 0));
    setActualCash(cashShift.actualCash != null ? String(cashShift.actualCash) : '');
    setShiftComment(cashShift.closeComment ?? '');
  }, [cashShift.id, cashShift.status, cashShift.openingCash, cashShift.actualCash, cashShift.closeComment]);

  const submitOpenShift = () => {
    if (openCashShift(openingCash)) {
      setActualCash('');
      setShiftComment('');
    }
  };

  const submitCloseShift = () => {
    if (closeCashShift(actualCash, shiftComment)) {
      setPaymentActionDraft(null);
    }
  };

  const submitPaymentAction = () => {
    if (!paymentActionDraft) return;
    if (paymentActionDraft.action === 'cancel') {
      cancelPayment(paymentActionDraft.paymentId, paymentActionDraft.reason);
    } else {
      refundPayment(paymentActionDraft.paymentId, paymentActionDraft.amount, paymentActionDraft.reason, paymentActionDraft.method);
    }
    setPaymentActionDraft(null);
  };

  return (
    <div className="page-grid">
      <PageTitle eyebrow="Оплати / Каса" title="Каса, термінал, безготівка і контроль зміни" text="Кожен платіж прив'язаний до співробітника, способу оплати і касової зміни." />
      <section className="stats-grid">
        <Metric icon={<Banknote />} label="Готівка в касі" value={money(expectedCash)} hint={cashShift.status === 'Відкрита' ? `зміна ${cashShift.id}` : 'зміна закрита'} />
        <Metric icon={<CheckCircle2 />} label="Термінал" value={money(cashShift.cardIncome)} hint="проведені карткою" />
        <Metric icon={<Archive />} label="Безготівка" value={money(cashShift.bankIncome)} hint="IBAN / банк" />
        <Metric icon={<PackageCheck />} label="Повернення" value={money(cashShift.cashExpense)} hint="готівкові повернення" />
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Касова зміна</h2><span>{cashShift.status}</span></div>
        <div className="info-grid">
          <Info label="Відкрив" value={cashShift.openedBy || '—'} />
          <Info label="Відкрито" value={cashShift.openedAt || '—'} />
          <Info label="Закрив" value={cashShift.closedBy || '—'} />
          <Info label="Закрито" value={cashShift.closedAt || '—'} />
          <Info label="Стартова готівка" value={money(cashShift.openingCash || 0)} />
          <Info label="Очікуваний залишок" value={money(expectedCash)} />
        </div>
        {cashShift.status === 'Закрита' ? (
          <div className="table purchase-detail-form">
            <label>
              Початковий залишок
              <input type="number" min={0} value={openingCash} onChange={(event) => setOpeningCash(event.target.value)} />
            </label>
            <div className="action-row">
              <button type="button" className="submit-button" onClick={submitOpenShift}>Відкрити зміну</button>
            </div>
          </div>
        ) : (
          <div className="table purchase-detail-form">
            <label>
              Фактичний залишок
              <input type="number" min={0} value={actualCash} onChange={(event) => setActualCash(event.target.value)} />
            </label>
            <label>
              Коментар до закриття
              <input value={shiftComment} onChange={(event) => setShiftComment(event.target.value)} placeholder="Обов'язково при розбіжності" />
            </label>
            <label>
              Різниця
              <input value={actualCash.trim() ? money(closeDifferencePreview) : '—'} readOnly />
            </label>
            <div className="action-row">
              <button type="button" className="submit-button" onClick={submitCloseShift}>Закрити зміну</button>
            </div>
          </div>
        )}
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Контроль грошей</h2><span>{payments.length} операцій</span></div>
        <div className="info-grid">
          <Info label="Готівка на початок" value={money(cashShift.openingCash || 0)} />
          <Info label="Готівкові оплати" value={money(cashShift.cashIncome)} />
          <Info label="Повернення готівкою" value={money(cashShift.cashExpense)} />
          <Info label="Термінал" value={money(cashShift.cardIncome)} />
          <Info label="Безготівка / IBAN" value={money(cashShift.bankIncome)} />
          <Info label="Спірні / не підтверджені" value={String(unconfirmedPayments.length)} />
        </div>
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Журнал каси</h2><span>{payments.length} операцій</span></div>
        <div className="table payments-table">
          <div className="table-row table-head"><span>Дата</span><span>Клієнт</span><span>Документ</span><span>Сума</span><span>Спосіб</span><span>Статус</span><span>Співробітник</span></div>
          {payments.map((payment) => {
            const canReverse = payment.amount > 0 && paymentCountsAsApplied(payment) && !['Скасовано', 'Повернення'].includes(payment.status ?? '');
            return (
              <div className="table-row" key={payment.id}>
                <span>{payment.date}</span>
                <span>{payment.client}</span>
                <span>{payment.document}</span>
                <span>{money(payment.amount)}</span>
                <span>{payment.method}<small>{payment.cashShiftId || payment.transactionNo}</small></span>
                <span>
                  {paymentProcessingLabel(payment)}
                  {paymentNeedsConfirmation(payment) && ['Бухгалтер', 'Менеджер', 'Адміністратор'].includes(activeUser.role) ? <button type="button" onClick={() => confirmBankPayment(payment.id)}>Підтвердити</button> : null}
                  {canReverse ? <button type="button" onClick={() => setPaymentActionDraft({ paymentId: payment.id, action: 'cancel', amount: String(Math.abs(payment.amount)), reason: 'Скасування платежу', method: payment.method })}>Скасувати</button> : null}
                  {canReverse ? <button type="button" onClick={() => setPaymentActionDraft({ paymentId: payment.id, action: 'refund', amount: String(Math.abs(payment.amount)), reason: 'Повернення коштів', method: payment.method })}>Повернення</button> : null}
                </span>
                <span>{payment.acceptedBy}</span>
              </div>
            );
          })}
        </div>
        {paymentActionDraft && (
          <div className="panel-subsection">
            <div className="panel-heading"><h2>{paymentActionDraft.action === 'cancel' ? 'Скасувати платіж' : 'Повернути кошти'}</h2><span>{paymentActionDraft.paymentId}</span></div>
            <div className="table purchase-detail-form">
              {paymentActionDraft.action === 'refund' && (
                <label>
                  Сума
                  <input type="number" min={0} value={paymentActionDraft.amount} onChange={(event) => setPaymentActionDraft((current) => current ? { ...current, amount: event.target.value } : current)} />
                </label>
              )}
              <label>
                Спосіб
                <select value={paymentActionDraft.method} onChange={(event) => setPaymentActionDraft((current) => current ? { ...current, method: event.target.value as PaymentMethod } : current)}>
                  <option value="Готівка">Готівка</option>
                  <option value="Картка">Термінал</option>
                  <option value="Безготівка">Безготівка / IBAN</option>
                </select>
              </label>
              <label>
                Причина
                <input value={paymentActionDraft.reason} onChange={(event) => setPaymentActionDraft((current) => current ? { ...current, reason: event.target.value } : current)} />
              </label>
            </div>
            <div className="action-row">
              <button type="button" className="submit-button" onClick={submitPaymentAction}>Підтвердити</button>
              <button type="button" onClick={() => setPaymentActionDraft(null)}>Скасувати</button>
            </div>
          </div>
        )}
      </section>
    </div>
  );
}

function BankImportPage({
  items,
  draft,
  allCandidates,
  bankAccounts,
  onImport,
  onConfirmReview,
  onFinalizeMapping,
}: {
  items: BankImportItem[];
  draft: BankImportDraft | null;
  allCandidates: BankImportCandidate[];
  bankAccounts: BankAccountRecord[];
  onImport: (file: File) => Promise<void>;
  onConfirmReview: (itemId: string, candidateKey: string) => void;
  onFinalizeMapping: (mapping: Partial<Record<BankImportField, string>>) => void;
}) {
  const importInputRef = useRef<HTMLInputElement | null>(null);
  const [reviewSelection, setReviewSelection] = useState<Record<string, string>>({});
  const [draftMapping, setDraftMapping] = useState<Partial<Record<BankImportField, string>>>({});
  const [statusFilter, setStatusFilter] = useState<'all' | 'matched' | 'review' | 'unmatched'>('all');
  const [isDragActive, setIsDragActive] = useState(false);
  const autoMatched = items.filter((item) => item.status === 'matched' && item.matchedBy === 'auto');
  const reviewItems = items.filter((item) => item.status === 'review');
  const unmatchedItems = items.filter((item) => item.status === 'unmatched');

  useEffect(() => {
    setDraftMapping(draft?.mapping ?? {});
  }, [draft]);

  async function handleFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      await onImport(file);
      setReviewSelection({});
    } catch (error) {
      window.alert(error instanceof Error ? error.message : 'Не вдалося імпортувати виписку.');
    } finally {
      event.target.value = '';
    }
  }

  async function handleDroppedFile(file: File) {
    try {
      await onImport(file);
      setReviewSelection({});
    } catch (error) {
      window.alert(error instanceof Error ? error.message : 'Не вдалося імпортувати виписку.');
    }
  }

  return (
    <div className="page-grid">
      <PageTitle eyebrow="Банк" title="Імпорт банківської виписки" text="Автоматичне підтвердження тільки для точних збігів. Усе сумнівне переходить на ручну перевірку." />
      <section className="panel executive-panel finance-center-filters">
        <div
          className="empty-state"
          onDragOver={(event) => {
            event.preventDefault();
            setIsDragActive(true);
          }}
          onDragLeave={() => setIsDragActive(false)}
          onDrop={async (event) => {
            event.preventDefault();
            setIsDragActive(false);
            const file = event.dataTransfer.files?.[0];
            if (file) await handleDroppedFile(file);
          }}
          style={isDragActive ? { borderColor: '#2563eb', color: '#2563eb' } : undefined}
        >
          Перетягніть CSV / XLSX / DBF у це поле або завантажте файл кнопкою нижче
        </div>
        <div className="quick-actions">
          <input ref={importInputRef} type="file" accept=".xlsx,.csv,.dbf" onChange={handleFileChange} style={{ display: 'none' }} />
          <button type="button" onClick={() => importInputRef.current?.click()}>Імпортувати виписку</button>
          <button type="button" className={statusFilter === 'all' ? 'primary' : ''} onClick={() => setStatusFilter('all')}>Усі</button>
          <button type="button" className={statusFilter === 'matched' ? 'primary' : ''} onClick={() => setStatusFilter('matched')}>Автоматично знайдено</button>
          <button type="button" className={statusFilter === 'review' ? 'primary' : ''} onClick={() => setStatusFilter('review')}>Требує перевірки</button>
          <button type="button" className={statusFilter === 'unmatched' ? 'primary' : ''} onClick={() => setStatusFilter('unmatched')}>Нерозпізнані</button>
        </div>
        <div className="executive-list">
          {bankAccounts.map((account) => (
            <div key={account.id} className="executive-list-row">
              <span>{account.bankName} · {account.legalType} · {account.currency}</span>
              <strong>{account.iban}</strong>
            </div>
          ))}
        </div>
      </section>

      {draft && (
        <section className="panel">
          <div className="panel-heading"><h2>Ручне зіставлення колонок</h2><span>{draft.detectedBank}</span></div>
          <div className="table">
            {(['date', 'amount', 'currency', 'payer', 'taxId', 'purpose', 'documentRef', 'direction', 'bankAccount'] as BankImportField[]).map((field) => (
              <label key={field}>
                {field}
                <select value={draftMapping[field] ?? ''} onChange={(event) => setDraftMapping((current) => ({ ...current, [field]: event.target.value || undefined }))}>
                  <option value="">Не вибрано</option>
                  {draft.headers.map((header) => <option key={`${field}-${header}`} value={header}>{header}</option>)}
                </select>
              </label>
            ))}
          </div>
          <div className="action-row">
            <button type="button" className="submit-button" onClick={() => onFinalizeMapping(draftMapping)}>Застосувати мапу</button>
          </div>
        </section>
      )}

      {(statusFilter === 'all' || statusFilter === 'matched') && (
      <section className="panel">
        <div className="panel-heading"><h2>Автоматично знайдено</h2><span>{autoMatched.length}</span></div>
        <div className="table">
          <div className="table-row table-head"><span>Дата</span><span>Сума</span><span>Платник</span><span>Рахунок</span><span>Призначення</span><span>Збіг</span></div>
          {autoMatched.map((item) => (
            <div key={item.id} className="table-row">
              <span>{item.row.date}</span>
              <span>{money(item.row.amount)}</span>
              <span>{item.row.payer}</span>
              <span>{item.row.bankAccount || '—'}</span>
              <span>{item.row.purpose || item.row.documentRef}</span>
              <span>{item.candidates.find((candidate) => candidate.key === item.matchedCandidateKey)?.label ?? item.reason}</span>
            </div>
          ))}
        </div>
        {autoMatched.length === 0 && <div className="empty-state">Автоматичних збігів поки немає.</div>}
      </section>
      )}

      {(statusFilter === 'all' || statusFilter === 'review') && (
      <section className="panel">
        <div className="panel-heading"><h2>Требує перевірки</h2><span>{reviewItems.length}</span></div>
        <div className="table">
          <div className="table-row table-head"><span>Дата</span><span>Сума</span><span>Платник</span><span>Рахунок</span><span>Призначення</span><span>Привʼязка</span></div>
          {reviewItems.map((item) => (
            <div key={item.id} className="table-row">
              <span>{item.row.date}</span>
              <span>{money(item.row.amount)}</span>
              <span>{item.row.payer}</span>
              <span>{item.row.bankAccount || '—'}</span>
              <span>{item.row.purpose || item.row.documentRef}</span>
              <span>
                <select value={reviewSelection[item.id] ?? ''} onChange={(event) => setReviewSelection((current) => ({ ...current, [item.id]: event.target.value }))}>
                  <option value="">Оберіть документ</option>
                  {item.candidates.map((candidate) => <option key={candidate.key} value={candidate.key}>{candidate.label} · {money(candidate.amount)}</option>)}
                </select>
                <button type="button" onClick={() => onConfirmReview(item.id, reviewSelection[item.id] ?? '')} disabled={!reviewSelection[item.id]}>Підтвердити</button>
              </span>
            </div>
          ))}
        </div>
        {reviewItems.length === 0 && <div className="empty-state">Немає платежів, що потребують перевірки.</div>}
      </section>
      )}

      {(statusFilter === 'all' || statusFilter === 'unmatched') && (
      <section className="panel">
        <div className="panel-heading"><h2>Нерозпізнані платежі</h2><span>{unmatchedItems.length}</span></div>
        <div className="table">
          <div className="table-row table-head"><span>Дата</span><span>Сума</span><span>Платник</span><span>Рахунок</span><span>ЄДРПОУ / ІПН</span><span>Призначення</span></div>
          {unmatchedItems.map((item) => (
            <div key={item.id} className="table-row">
              <span>{item.row.date}</span>
              <span>{money(item.row.amount)}</span>
              <span>{item.row.payer}</span>
              <span>{item.row.bankAccount || '—'}</span>
              <span>{item.row.taxId || '—'}</span>
              <span>
                {item.row.purpose || item.row.documentRef || '—'}
                <select value={reviewSelection[item.id] ?? ''} onChange={(event) => setReviewSelection((current) => ({ ...current, [item.id]: event.target.value }))}>
                  <option value="">Вручну привʼязати</option>
                  {allCandidates.map((candidate) => <option key={`${item.id}-${candidate.key}`} value={candidate.key}>{candidate.label} · {money(candidate.amount)}</option>)}
                </select>
                <button type="button" onClick={() => onConfirmReview(item.id, reviewSelection[item.id] ?? '')} disabled={!reviewSelection[item.id]}>Підтвердити</button>
              </span>
            </div>
          ))}
        </div>
        {unmatchedItems.length === 0 && <div className="empty-state">Нерозпізнаних платежів зараз немає.</div>}
      </section>
      )}
    </div>
  );
}

function ReturnsPage({ sales, orders, products, productName }: { sales: Sale[]; orders: ServiceOrder[]; products: Product[]; productName: (id: string) => string }) {
  const saleReturns = sales.flatMap((sale) => sale.returns.map((item) => ({ id: item.id, source: sale.id, client: sale.client, product: productName(item.productId), qty: item.qty, amount: item.refund, reason: item.reason, decision: item.destination })));
  const serviceReturns = orders.flatMap((order) => order.parts.filter((part) => part.status === 'Повернення').map((part) => ({ id: part.id, source: order.id, client: order.client, product: productName(part.productId), qty: part.qty, amount: part.price * part.qty, reason: 'Деталь не підійшла або знята з ремонту', decision: 'Сервісне повернення' })));
  const rows = [...saleReturns, ...serviceReturns];
  const defectCandidates = products.filter((product) => product.installed > 0 || product.withEngineer > 0);

  return (
    <div className="page-grid">
      <PageTitle eyebrow="Повернення" title="Повернення по продажах і ремонтах" text="В одному місці видно причину повернення, суму, документ-джерело і рішення по товару: склад або брак." />
      <section className="stats-grid">
        <Metric icon={<PackageCheck />} label="Повернень" value={String(rows.length)} hint="продажі + сервіс" />
        <Metric icon={<Banknote />} label="Сума повернень" value={money(rows.reduce((sum, item) => sum + item.amount, 0))} hint="до контролю бухгалтером" />
        <Metric icon={<Archive />} label="Кандидати в брак" value={String(defectCandidates.length)} hint="потребують рішення" />
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Журнал повернень</h2><span>{rows.length} записів</span></div>
        <div className="table returns-table">
          <div className="table-row table-head"><span>Документ</span><span>Клієнт</span><span>Товар</span><span>К-сть</span><span>Сума</span><span>Рішення</span></div>
          {rows.map((item) => (
            <div className="table-row" key={item.id}>
              <span>{item.source}</span>
              <span>{item.client}</span>
              <span>{item.product}</span>
              <span>{item.qty}</span>
              <span>{money(item.amount)}</span>
              <span>{item.decision}</span>
            </div>
          ))}
        </div>
        {rows.length === 0 && <div className="empty-state">Повернень поки немає.</div>}
      </section>
    </div>
  );
}

function DocumentsPage({ documents, templates }: { documents: PrintDocument[]; templates: Array<{ kind: DocumentKind; description: string; lockedRule: string }> }) {
  const sampleOrder = initialOrders[1] ?? initialOrders[0];
  const sampleTotals = orderTotals(sampleOrder);
  const sampleVat = documentVatTotals(sampleTotals.total);
  const sampleClient = clientDetails(sampleOrder.client, sampleOrder.phone);
  const sampleProductName = (productId: string) => initialProducts.find((product) => product.id === productId)?.name ?? productId;
  return (
    <div className="page-grid">
      <PageTitle eyebrow="Документи" title="PDF, друк і юридичний архів" text="Квитанції, акти, чеки, накладні та повернення створюються з карток замовлень, продажів і закупівель. Кожен документ має версію, автора, PDF-шлях і звʼязок з документом-джерелом." />
      <section className="stats-grid">
        <Metric icon={<ClipboardList />} label="Документів" value={String(documents.length)} hint="збережені PDF" />
        <Metric icon={<ShieldCheck />} label="Шаблонів" value={String(templates.length)} hint="для ремонту, продажу, складу" />
        <Metric icon={<History />} label="Версійність" value="Увімкнено" hint="підписані документи блокуються" />
        <Metric icon={<CheckCircle2 />} label="Автозвʼязок" value="Є" hint="замовлення, продаж, закупка" />
      </section>
      <section className="content-split">
        <div className="panel document-preview">
          <div className="panel-heading"><h2>Рахунок на оплату</h2><span>офіційний бланк</span></div>
          <div className="official-document">
            <h3>РАХУНОК НА ОПЛАТУ № РХ-1001</h3>
            <p>від {today}</p>
            <div className="details-grid">
              <Info label="Постачальник" value={`${companyRequisites.name}, ЄДРПОУ ${companyRequisites.edrpou}, ${companyRequisites.iban}, ${companyRequisites.bank}, ${companySettings.mainEmail}`} />
              <Info label="Покупець" value={`${sampleClient.name}, ЄДРПОУ ${sampleClient.edrpou}, ${sampleClient.address}`} />
              <Info label="Договір" value={sampleClient.contract} />
              <Info label="Сума прописью" value={amountInWordsUA(sampleVat.amountGross)} />
            </div>
            <div className="table documents-table">
              <div className="table-row table-head"><span>Найменування</span><span>К-сть</span><span>Ціна</span><span>Сума</span></div>
              {sampleOrder.works.map((work) => <div className="table-row" key={work.name}><span>{work.name}</span><span>{work.qty ?? 1}</span><span>{money(work.price)}</span><span>{money(work.price * (work.qty ?? 1))}</span></div>)}
              {sampleOrder.parts.map((part) => <div className="table-row" key={part.id}><span>{sampleProductName(part.productId)}</span><span>{part.qty}</span><span>{money(part.price)}</span><span>{money(part.price * part.qty)}</span></div>)}
            </div>
            <div className="details-grid">
              <Info label="Сума без ПДВ" value={money(sampleVat.amountNet)} />
              <Info label="ПДВ 20%" value={money(sampleVat.amountVat)} />
              <Info label="Усього з ПДВ" value={money(sampleVat.amountGross)} />
            </div>
          </div>
        </div>
        <div className="panel document-preview">
          <div className="panel-heading"><h2>Акт надання послуг</h2><span>роботи і комплектуючі окремо</span></div>
          <div className="official-document">
            <h3>АКТ НАДАННЯ ПОСЛУГ № АНП-1001</h3>
            <p>місце складання: м. Київ · дата {today}</p>
            <p>{companyRequisites.name}, в особі {companyRequisites.signatory}, email {companySettings.mainEmail}, та {sampleClient.name} склали цей акт за договором: {sampleClient.contract}.</p>
            <h4>Виконані роботи / послуги</h4>
            <div className="table documents-table">
              {sampleOrder.works.map((work) => <div className="table-row" key={work.name}><span>{work.name}</span><span>{work.qty ?? 1}</span><span>{money(work.price * (work.qty ?? 1))}</span></div>)}
            </div>
            <h4>Використані комплектуючі</h4>
            <div className="table documents-table">
              {sampleOrder.parts.map((part) => <div className="table-row" key={part.id}><span>{sampleProductName(part.productId)}</span><span>{part.qty}</span><span>{money(part.price * part.qty)}</span></div>)}
              {sampleOrder.parts.length === 0 && <div className="empty-state">Комплектуючі не використовувались.</div>}
            </div>
            <div className="details-grid">
              <Info label="Сума без ПДВ" value={money(sampleVat.amountNet)} />
              <Info label="ПДВ 20%" value={money(sampleVat.amountVat)} />
              <Info label="Усього з ПДВ" value={money(sampleVat.amountGross)} />
              <Info label="Претензії" value="Сторони претензій щодо обсягу, якості та строків надання послуг не мають." />
              <Info label="Реквізити виконавця" value={`${companyRequisites.name}, ЄДРПОУ ${companyRequisites.edrpou}, ${companyRequisites.address}, ${companySettings.mainEmail}`} />
              <Info label="Реквізити замовника" value={`${sampleClient.name}, ЄДРПОУ ${sampleClient.edrpou}, ${sampleClient.address}`} />
            </div>
            <p>Підписи сторін: Виконавець ____________ / Замовник ____________</p>
          </div>
        </div>
      </section>
      <section className="content-split">
        <div className="panel document-preview">
          <div className="panel-heading"><h2>Видаткова накладна</h2><span>тільки товар / склад</span></div>
          <div className="official-document">
            <h3>ВИДАТКОВА НАКЛАДНА № ВН-1001</h3>
            <p>дата {today} · місце складання: м. Київ</p>
            <Info label="Постачальник" value={`${companyRequisites.name}, ЄДРПОУ ${companyRequisites.edrpou}, ${companySettings.mainEmail}`} />
            <Info label="Покупець" value={`${sampleClient.name}, ${sampleClient.contact}`} />
            <Info label="Договір" value={sampleClient.contract} />
            <div className="table documents-table">
              <div className="table-row table-head"><span>Товар</span><span>К-сть</span><span>Ціна без ПДВ</span><span>Сума без ПДВ</span></div>
              {sampleOrder.parts.map((part) => {
                const lineNet = vatFromGross(part.price).net;
                return <div className="table-row" key={part.id}><span>{sampleProductName(part.productId)}</span><span>{part.qty}</span><span>{money(lineNet)}</span><span>{money(lineNet * part.qty)}</span></div>;
              })}
            </div>
            <div className="details-grid">
              <Info label="Разом" value={money(sampleVat.amountNet)} />
              <Info label="ПДВ" value={money(sampleVat.amountVat)} />
              <Info label="Усього з ПДВ" value={money(sampleVat.amountGross)} />
              <Info label="Відпустив / Отримав" value="____________ / ____________" />
            </div>
          </div>
        </div>
        <div className="panel document-preview">
          <div className="panel-heading"><h2>Сервісні документи</h2><span>наряд, техстан, гарантія</span></div>
          <div className="task-list">
            <Task icon={<ClipboardList />} title="Наряд на ремонт" text={`№ НР-1001 · клієнт ${sampleOrder.client} · пристрій ${sampleOrder.device} · серійний номер ${sampleOrder.serial} · несправність: ${sampleOrder.issue} · підписи виконавця, менеджера і клієнта.`} />
            <Task icon={<ClipboardList />} title="Швидкий наряд" text="Для картриджів: клієнт, картридж, вид роботи, результат, підписи виконавця і клієнта." />
            <Task icon={<X />} title="Акт технічного стану" text="Комісія, обʼєкт, серійний номер, дефекти, висновок: не підлягає ремонту, підписи членів комісії." />
            <Task icon={<ShieldCheck />} title="Гарантійний талон" text={`Пристрій ${sampleOrder.device}, серійний номер ${sampleOrder.serial}, дата ${today}, термін гарантії, умови, підпис / печатка.`} />
            <Task icon={<Archive />} title="BAS" text="У кожному документі є дія Передати в BAS. CRM передає клієнта, товари / послуги, кількість, суми, ПДВ, номер і дату. Офіційна ПН і бухгалтерські документи формуються в BAS." />
          </div>
        </div>
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Реєстр документів</h2><span>{documents.length} записів</span></div>
        <div className="table documents-table">
          <div className="table-row table-head"><span>Номер</span><span>Тип</span><span>Джерело</span><span>Клієнт / постачальник</span><span>Статус</span><span>PDF / BAS</span></div>
          {documents.map((document) => (
            <div className="table-row" key={document.id}>
              <span>{document.number}<small>версія {document.version} · {document.createdAt}</small></span>
              <span>{document.kind}</span>
              <span>{document.entityId}<small>{document.entityType}</small></span>
              <span>{document.clientOrSupplier}<small>{document.createdBy}</small></span>
              <span>{document.status}</span>
              <span>
                {document.pdfPath}
                <small>{document.amountGross ? `без ПДВ ${money(document.amountNet ?? 0)} · ПДВ ${money(document.amountVat ?? 0)} · з ПДВ ${money(document.amountGross)}` : 'суми не задані'}</small>
                <button type="button">Передати в BAS</button>
                {document.status === 'Чернетка' || document.status === 'Готово' || document.status === 'Очікує оплати' || document.status === 'Готово до видачі' ? <button type="button">Редагувати чернетку</button> : null}
                {document.status !== 'Підписано' ? <button type="button">Пересоздати</button> : null}
                {document.status === 'Чернетка' ? <button type="button">Видалити чернетку</button> : null}
              </span>
            </div>
          ))}
        </div>
        {documents.length === 0 && <div className="empty-state">Документи ще не створені. Реєстр не створює їх вручну: відкрийте замовлення, продаж або закупку і натисніть потрібну кнопку процесу.</div>}
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Шаблони друкованих форм</h2><span>логотип, реквізити, підписи, печатка</span></div>
        <div className="task-list">
          {templates.map((template) => (
            <Task key={template.kind} icon={<ClipboardList />} title={template.kind} text={`${template.description} ${template.lockedRule}`} />
          ))}
        </div>
      </section>
    </div>
  );
}

function TaxInvoicesPage({ invoices, orders, createTaxInvoiceForOrder, registerTaxInvoice }: { invoices: TaxInvoice[]; orders: ServiceOrder[]; createTaxInvoiceForOrder: (orderId: string) => void; registerTaxInvoice: (orderId: string) => void }) {
  const created = invoices.filter((invoice) => invoice.status === 'Створено');
  const registered = invoices.filter((invoice) => invoice.status === 'Зареєстровано');
  const errors = invoices.filter((invoice) => invoice.status === 'Помилка');

  return (
    <div className="page-grid">
      <PageTitle eyebrow="Податкові накладні" title="ПН як окремий документ CRM" text="ПН створюється вручну після першої події, має snapshot і не змінюється після створення." />
      <section className="stats-grid">
        <Metric icon={<ClipboardList />} label="Створено" value={String(created.length)} hint="чекають реєстрації" />
        <Metric icon={<CheckCircle2 />} label="Зареєстровано" value={String(registered.length)} hint="успішно завершені" />
        <Metric icon={<X />} label="Помилки" value={String(errors.length)} hint="потребують уваги" />
        <Metric icon={<Banknote />} label="ПДВ до контролю" value={money(invoices.reduce((sum, item) => sum + item.vatAmount, 0))} hint="за чернетками ПН" />
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Створені ПН</h2><span>податкові документи зі snapshot</span></div>
        <div className="task-list">
          {created.map((invoice) => {
            const snapshot = parseTaxInvoiceSnapshot(invoice.snapshot);
            return (
              <Task
                key={`draft-${invoice.id}`}
                icon={<ClipboardList />}
                title={`${invoice.number} · ${invoice.client}`}
                text={`Подія: ${snapshot?.eventType === 'payment' ? 'оплата' : 'акт'}, дата ${invoice.date}, замовлення ${invoice.orderId}, ІПН клієнта ${snapshot?.clientTaxId || 'не вказано'}, ПДВ ${money(invoice.vatAmount)}.`}
              />
            );
          })}
          {created.length === 0 && <div className="empty-state">Створених ПН для реєстрації зараз немає.</div>}
        </div>
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Реєстр ПН</h2><span>звʼязок з оплатою або актом</span></div>
        <div className="table tax-table">
          <div className="table-row table-head"><span>ПН</span><span>Замовлення</span><span>Клієнт</span><span>Подія</span><span>Сума / ПДВ</span><span>Статус</span><span>Дія</span></div>
          {invoices.map((invoice) => {
            const snapshot = parseTaxInvoiceSnapshot(invoice.snapshot);
            return (
              <div className="table-row" key={invoice.id}>
                <span>{invoice.number}<small>{invoice.createdAt}</small></span>
                <span>{invoice.orderId}<small>{invoice.date}</small></span>
                <span>{invoice.client}</span>
                <span>{invoice.eventDate}<small>{snapshot?.eventType === 'payment' ? 'перша подія: оплата' : 'перша подія: акт'}</small></span>
                <span>{money(invoice.amount)}<small>ПДВ {money(invoice.vatAmount)}</small></span>
                <span>{invoice.status}</span>
                <span>{invoice.status !== 'Зареєстровано' ? <button type="button" onClick={() => registerTaxInvoice(invoice.orderId)}>Позначити як зареєстровано</button> : '—'}</span>
              </div>
            );
          })}
        </div>
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Замовлення з ПДВ</h2><span>ручне створення ПН після події</span></div>
        <div className="table tax-table">
          <div className="table-row table-head"><span>Замовлення</span><span>Клієнт</span><span>Сума</span><span>Статус ПН</span><span>Дія</span></div>
          {orders.filter((order) => order.legalEntity || Boolean(order.vatStatus)).map((order) => (
            <div className="table-row" key={`tax-order-${order.id}`}>
              {(() => {
                const orderInvoice = [...invoices]
                  .filter((invoice) => invoice.orderId === order.id)
                  .sort((a, b) => (parseDateTime(b.createdAt)?.getTime() ?? 0) - (parseDateTime(a.createdAt)?.getTime() ?? 0))[0];
                return (
                  <>
              <span>{order.id}</span>
              <span>{order.client}</span>
              <span>{money(orderTotals(order).total)}</span>
              <span>{orderInvoice ? `${orderInvoice.number} · ${orderInvoice.status}` : (order.vatStatus ?? '—')}</span>
              <span>
                {!orderInvoice
                  ? <button type="button" onClick={() => createTaxInvoiceForOrder(order.id)}>Створити ПН</button>
                  : orderInvoice.status === 'Створено'
                    ? <button type="button" onClick={() => registerTaxInvoice(order.id)}>Позначити як зареєстровано</button>
                    : '—'}
              </span>
                  </>
                );
              })()}
            </div>
          ))}
        </div>
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Правила</h2><span>обовʼязково</span></div>
        <div className="task-list">
          <Task icon={<CheckCircle2 />} title="Перша подія" text="ПН створюється вручну після підтвердженої оплати або підписаного акта. Автоматичного створення немає." />
          <Task icon={<Banknote />} title="Snapshot" text="Після створення ПН більше не перераховується. Реквізити, рядки і суми заморожені в snapshot." />
          <Task icon={<Bell />} title="Видачу не блокує" text="ПН не зупиняє видачу замовлення, але залишається окремим етапом для менеджера і бухгалтера." />
        </div>
      </section>
    </div>
  );
}

function BasExchangePage({ items, mappings, actionLogs }: { items: BasExchangeItem[]; mappings: BasMapping[]; actionLogs: ActionLog[] }) {
  const byStatus = (status: BasExchangeStatus) => items.filter((item) => item.exchangeStatus === status).length;
  const exportRows = [
    { name: 'Контрагенти', value: items.filter((item) => item.platformEntityType === 'counterparty').length },
    { name: 'Акти', value: items.filter((item) => item.platformEntityType === 'act').length },
    { name: 'Оплати', value: items.filter((item) => item.platformEntityType === 'payment').length },
    { name: 'Продажі', value: items.filter((item) => item.platformEntityType === 'sale').length },
    { name: 'Закупки', value: items.filter((item) => item.platformEntityType === 'purchase').length },
    { name: 'ПН', value: items.filter((item) => item.platformEntityType === 'tax_invoice').length },
  ];
  const duplicateKeys = items
    .map((item) => `${item.platformEntityType}:${item.platformEntityId}`)
    .filter((key, index, list) => list.indexOf(key) !== index);
  const sections: Array<{ title: string; types: BasObjectType[] }> = [
    { title: 'Контрагенти', types: ['counterparty'] },
    { title: 'Акти', types: ['act'] },
    { title: 'Оплати', types: ['payment'] },
    { title: 'Податкові накладні', types: ['tax_invoice'] },
    { title: 'Закупки', types: ['purchase'] },
    { title: 'Долги', types: ['payment', 'act'] },
  ];

  return (
    <div className="page-grid">
      <PageTitle eyebrow="Обмін з BAS" title="Операційна платформа -> бухгалтерія" text="CRM є первинним джерелом заказів, ремонту, складу і оплат. BAS є первинним джерелом бухгалтерського обліку, звітності і ПДВ." />
      <section className="stats-grid">
        <Metric icon={<ClipboardList />} label="New" value={String(byStatus('new'))} hint="потрібно поставити в чергу" />
        <Metric icon={<History />} label="Queued" value={String(byStatus('queued'))} hint="готово до выгрузки" />
        <Metric icon={<Truck />} label="Sent" value={String(byStatus('sent'))} hint="відправлено" />
        <Metric icon={<CheckCircle2 />} label="Imported" value={String(byStatus('imported'))} hint="прийнято BAS" />
        <Metric icon={<X />} label="Error" value={String(byStatus('error'))} hint="потрібне виправлення" />
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Черговість впровадження</h2><span>спочатку робота сервісу, потім BAS</span></div>
        <div className="task-list">
          {basImplementationPhases.map((phase) => (
            <Task
              key={phase.phase}
              icon={phase.done ? <CheckCircle2 /> : phase.priority === 'Не робити на старті' ? <X /> : <History />}
              title={`${phase.phase} · ${phase.title}`}
              text={`${phase.priority}. ${phase.result} Склад: ${phase.items.join(', ')}.`}
            />
          ))}
        </div>
      </section>
      <section className="quick-actions">
        <button type="button">Сформувати ПН</button>
        <button type="button">Вигрузити в BAS</button>
        <button type="button">Вигрузити всі нові</button>
        <button type="button">Підтвердити оплату</button>
        <button type="button">Повторити відправку</button>
        <button type="button">Відкрити помилку</button>
        <button type="button">Excel / CSV / JSON</button>
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Контроль унікальності</h2><span>platform_entity_type + platform_entity_id</span></div>
        <div className="task-list">
          <Task icon={<ShieldCheck />} title="Повторне створення заборонене" text="Один і той самий документ не повинен створюватися в BAS повторно. Ключ унікальності: platform_entity_type + platform_entity_id." />
          <Task icon={duplicateKeys.length ? <X /> : <CheckCircle2 />} title="Дублі в черзі" text={duplicateKeys.length ? duplicateKeys.join(', ') : 'Дублів не знайдено.'} />
          <Task icon={<Archive />} title="Payload" text="Кожен документ має payload_json, який можна повторно відправити після виправлення помилки." />
        </div>
      </section>
      <section className="content-split">
        {sections.map((section) => {
          const rows = items.filter((item) => section.types.includes(item.platformEntityType));
          return (
            <div className="panel" key={section.title}>
              <div className="panel-heading"><h2>{section.title}</h2><span>{rows.length} документів</span></div>
              <div className="task-list">
                {rows.slice(0, 3).map((item) => <Task key={`${section.title}-${item.id}`} icon={<ClipboardList />} title={`${item.platformEntityId} · ${item.exchangeStatus}`} text={`${item.client} · ${money(item.amountGross)} · ПДВ ${money(item.amountVat)}${item.lastError ? ` · ${item.lastError}` : ''}`} />)}
                {rows.length === 0 && <div className="empty-state">Документів немає.</div>}
              </div>
            </div>
          );
        })}
      </section>
      <section className="content-split">
        <div className="panel">
          <div className="panel-heading"><h2>Що передається</h2><span>етап 1: файли, етап 2: API</span></div>
          <div className="task-list">
            {exportRows.map((row) => <Task key={row.name} icon={<Archive />} title={row.name} text={`${row.value} обʼєктів у поточній черзі або реєстрі обміну.`} />)}
          </div>
        </div>
        <div className="panel">
          <div className="panel-heading"><h2>Архітектура</h2><span>відповідальність систем</span></div>
          <div className="task-list">
            <Task icon={<LayoutDashboard />} title="Платформа" text="Первинне джерело: заказы, ремонти, склад, оплати, операційні статуси." />
            <Task icon={<Banknote />} title="BAS" text="Первинне джерело: бухгалтерський облік, регламентована звітність, ПДВ." />
            <Task icon={<ClipboardList />} title="Перший етап" text="Вигрузка Excel / CSV / JSON для бухгалтера." />
            <Task icon={<CheckCircle2 />} title="Другий етап" text="Автоматичний модуль обміну через API / файл-обмін після погодження формату." />
            <Task icon={<ShieldCheck />} title="Обмеження CRM" text="CRM не веде бухгалтерський облік і не формує офіційну податкову накладну. CRM готує документи та дані, BAS формує ПН і бухгалтерські документи." />
          </div>
        </div>
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Черга обміну</h2><span>{items.length} обʼєктів</span></div>
        <div className="table documents-table">
          <div className="table-row table-head"><span>Документ</span><span>Клієнт</span><span>Суми</span><span>Статус</span><span>BAS</span><span>Retry / Error</span></div>
          {items.map((item) => (
            <div className="table-row" key={item.id}>
              <span>{item.title}<small>{item.platformEntityType} · {item.platformEntityId} · створив {item.createdBy}</small></span>
              <span>{item.client}<small>заказ {item.orderId ?? 'немає'}</small></span>
              <span>{money(item.amountNet)}<small>ПДВ {money(item.amountVat)} · з ПДВ {money(item.amountGross)}</small></span>
              <span>{item.exchangeStatus}<small>sent {item.sentAt ?? 'ні'} · confirmed {item.confirmedAt ?? 'ні'}</small></span>
              <span>{item.basEntityId ?? 'немає'}<small>{item.payloadJson}</small></span>
              <span>{item.retryCount}<small>{item.lastError ?? 'немає'}</small></span>
            </div>
          ))}
        </div>
      </section>
      <section className="content-split">
        <div className="panel">
          <div className="panel-heading"><h2>Сопоставлення довідників</h2><span>{mappings.length} правил</span></div>
          <div className="table documents-table">
            <div className="table-row table-head"><span>Довідник</span><span>Платформа</span><span>BAS</span><span>Статус</span></div>
            {mappings.map((mapping) => (
              <div className="table-row" key={`${mapping.directory}-${mapping.platformValue}`}>
                <span>{mapping.directory}</span>
                <span>{mapping.platformValue}</span>
                <span>{mapping.basValue}</span>
                <span>{mapping.status}</span>
              </div>
            ))}
          </div>
        </div>
        <div className="panel">
          <div className="panel-heading"><h2>Логування обміну</h2><span>усі дії фіксуються</span></div>
          <div className="task-list">
            {items.slice(0, 5).map((item) => <Task key={`log-${item.id}`} icon={<History />} title={`${item.exchangeStatus} · ${item.platformEntityId}`} text={`${item.title}. ${item.lastError ? `Помилка: ${item.lastError}` : `Відповідальний: ${item.responsible}`}`} />)}
            {actionLogs.slice(0, 3).map((log) => <Task key={`action-${log.id}`} icon={<Bell />} title={`${log.action} · ${log.entity}`} text={`${log.date} · ${log.user}. ${log.comment}`} />)}
          </div>
        </div>
      </section>
    </div>
  );
}

function ProblemClientsPage({ orders, documents, setPage, setSelectedOrderId }: { orders: ServiceOrder[]; documents: PrintDocument[]; setPage: (page: Page) => void; setSelectedOrderId: (id: string) => void }) {
  const [managerFilter, setManagerFilter] = useState('Усі');
  const [engineerFilter, setEngineerFilter] = useState('Усі');
  const [typeFilter, setTypeFilter] = useState('Усі');
  const [ageFilter, setAgeFilter] = useState('Усі');
  const problems = buildServiceProblems(orders, documents);
  const managers = Array.from(new Set(problems.map((problem) => problem.manager)));
  const engineers = Array.from(new Set(problems.map((problem) => problem.engineer)));
  const types = Array.from(new Set(problems.map((problem) => problem.type)));
  const rows = problems.filter((problem) => {
    if (managerFilter !== 'Усі' && problem.manager !== managerFilter) return false;
    if (engineerFilter !== 'Усі' && problem.engineer !== engineerFilter) return false;
    if (typeFilter !== 'Усі' && problem.type !== typeFilter) return false;
    if (ageFilter === '1+ день' && problem.days < 1) return false;
    if (ageFilter === '3+ дні' && problem.days < 3) return false;
    if (ageFilter === '7+ днів' && problem.days < 7) return false;
    return true;
  });
  const critical = rows.filter((problem) => problem.level === 'Критично');
  const important = rows.filter((problem) => problem.level === 'Важливо');

  return (
    <div className="page-grid">
      <PageTitle eyebrow="Проблеми" title="Автоматичний контроль помилок сервісу" text="CRM сама аналізує заказы, документи, оплату, ПДВ і рух процесу. Користувач не шукає помилки вручну, а реагує на готовий список." />
      <section className="stats-grid">
        <Metric icon={<X />} label="Критично" value={String(critical.length)} hint="гроші / податки" />
        <Metric icon={<Bell />} label="Важливо" value={String(important.length)} hint="процеси" />
        <Metric icon={<History />} label="Усього проблем" value={String(rows.length)} hint="після фільтрів" />
        <Metric icon={<CheckCircle2 />} label="Автооновлення" value="5-10 хв" hint="умови перевіряються системою" />
      </section>
      <section className="quick-actions">
        <select value={managerFilter} onChange={(event) => setManagerFilter(event.target.value)} aria-label="Менеджер">
          <option>Усі</option>
          {managers.map((manager) => <option key={manager}>{manager}</option>)}
        </select>
        <select value={engineerFilter} onChange={(event) => setEngineerFilter(event.target.value)} aria-label="Інженер">
          <option>Усі</option>
          {engineers.map((engineer) => <option key={engineer}>{engineer}</option>)}
        </select>
        <select value={typeFilter} onChange={(event) => setTypeFilter(event.target.value)} aria-label="Тип проблеми">
          <option>Усі</option>
          {types.map((type) => <option key={type}>{type}</option>)}
        </select>
        <select value={ageFilter} onChange={(event) => setAgeFilter(event.target.value)} aria-label="Строк">
          <option>Усі</option>
          <option>1+ день</option>
          <option>3+ дні</option>
          <option>7+ днів</option>
        </select>
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Список проблем</h2><span>{rows.length} сигналів</span></div>
        <div className="table problem-table">
          <div className="table-row table-head"><span>Рівень</span><span>Заказ / клієнт</span><span>Проблема</span><span>Триває</span><span>Відповідальний</span><span>Дії</span></div>
          {rows.map((problem) => (
            <div className="table-row" key={problem.id}>
              <span>{problem.level}</span>
              <span>{problem.orderId}<small>{problem.client}</small></span>
              <span>{problem.type}<small>{problem.description}</small></span>
              <span>{problem.days} дн.</span>
              <span>{problem.responsible}<small>менеджер {problem.manager} · інженер {problem.engineer}</small></span>
              <span>
                <button type="button" onClick={() => { setSelectedOrderId(problem.orderId); setPage('orders'); }}>Відкрити заказ</button>
                {problem.target === 'payment' && <button type="button" onClick={() => { setSelectedOrderId(problem.orderId); setPage('orders'); }}>До оплати</button>}
                {problem.target === 'act' && <button type="button" onClick={() => setPage('documents')}>До акта</button>}
                {problem.target === 'document' && <button type="button" onClick={() => setPage('documents')}>До документа</button>}
              </span>
            </div>
          ))}
        </div>
        {rows.length === 0 && <div className="empty-state">За поточними фільтрами проблем немає.</div>}
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Правила контролю</h2><span>перевірка кожні 5-10 хв</span></div>
        <div className="task-list">
          <Task icon={<Banknote />} title="Критично" text="Гроші, акти, видаткові, НДС і борги піднімаються як критичні ризики." />
          <Task icon={<History />} title="Важливо" text="Готово, але не видано; заказ без руху; акт не повернений клієнтом." />
          <Task icon={<Bell />} title="Уведомлення" text="Проблеми видно на головній керівника, головній менеджера, у центрі повідомлень і в цьому розділі." />
        </div>
      </section>
    </div>
  );
}

function OrderUnitsPage({ units, activeUser }: { units: OrderUnit[]; activeUser: User }) {
  const [scanCode, setScanCode] = useState('');
  const visibleUnits = activeUser.role === 'Інженер' ? units.filter((unit) => unit.engineer === activeUser.name) : units;
  const scanned = visibleUnits.find((unit) => unit.code.toLowerCase() === scanCode.trim().toLowerCase() || unit.id.toLowerCase() === scanCode.trim().toLowerCase());
  const onShelfWithoutCell = visibleUnits.filter((unit) => unit.status === 'На полці' && !unit.locationCode);

  return (
    <div className="page-grid">
      <PageTitle eyebrow="Одиниці заказів" title="Маркування, сканування і рух по цеху" text="Кожен пристрій або картридж має свій ID, QR/штрихкод, статус, інженера і місце зберігання." />
      <section className="panel">
        <div className="panel-heading"><h2>Сканування</h2><span>QR / barcode</span></div>
        <label>
          Код одиниці
          <input value={scanCode} onChange={(event) => setScanCode(event.target.value)} placeholder="Наприклад BAR-1047-C2 або UNIT-1048-1" />
          <small>Сканування відкриває одиницю і показує доступні дії.</small>
        </label>
        {scanned && <div className="scan-result"><strong>{scanned.id} · {scanned.type}</strong><span>{scanned.client} · {scanned.status} · {scanned.locationCode ?? 'без полки'}</span></div>}
      </section>
      <section className="quick-actions">
        <button type="button">Створити ще один заказ для цього клієнта</button>
        <button type="button">Друк наклейки</button>
        <button type="button">Покласти на полку</button>
        <button type="button">Видати через сканування</button>
      </section>
      <section className="stats-grid">
        <Metric icon={<ClipboardList />} label="Одиниць" value={String(visibleUnits.length)} hint="пристрої + картриджі" />
        <Metric icon={<Wrench />} label="В ремонті" value={String(visibleUnits.filter((unit) => unit.status === 'В ремонті').length)} hint="цех" />
        <Metric icon={<Archive />} label="На полці" value={String(visibleUnits.filter((unit) => unit.status === 'На полці').length)} hint="готові або очікують" />
        <Metric icon={<X />} label="Без комірки" value={String(onShelfWithoutCell.length)} hint="заборонено правилами" />
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Таблиця одиниць</h2><span>{visibleUnits.length} записів</span></div>
        <div className="table units-table">
          <div className="table-row table-head"><span>ID</span><span>Заказ</span><span>Тип</span><span>Статус</span><span>Інженер</span><span>Місце</span><span>Код</span></div>
          {visibleUnits.map((unit) => (
            <div className="table-row" key={unit.id}>
              <span>{unit.id}<small>{unit.groupId}</small></span>
              <span>{unit.orderId}<small>{unit.client}</small></span>
              <span>{unit.type}</span>
              <span>{unit.status}</span>
              <span>{unit.engineer}</span>
              <span>{unit.locationCode ?? 'не задано'}</span>
              <span>{unit.code}</span>
            </div>
          ))}
        </div>
      </section>
    </div>
  );
}

function StoragePage({ orders, locations, movementLogs }: { orders: ServiceOrder[]; locations: WarehouseLocation[]; movementLogs: OrderMovementLog[] }) {
  const [search, setSearch] = useState('');
  const [zoneFilter, setZoneFilter] = useState<'ALL' | WarehouseZone>('ALL');
  const [rackFilter, setRackFilter] = useState('ALL');
  const [shelfFilter, setShelfFilter] = useState('ALL');
  const occupiedLocations = new Map(orders.filter((order) => order.locationCode).map((order) => [order.locationCode as string, order]));
  const filteredLocations = locations.filter((location) => {
    const matchesSearch = !search.trim() || [location.code, location.zone, location.rack, location.shelf].join(' ').toLowerCase().includes(search.trim().toLowerCase());
    const matchesZone = zoneFilter === 'ALL' || location.zone === zoneFilter;
    const matchesRack = rackFilter === 'ALL' || location.rack === rackFilter;
    const matchesShelf = shelfFilter === 'ALL' || location.shelf === shelfFilter;
    return matchesSearch && matchesZone && matchesRack && matchesShelf;
  });
  const free = filteredLocations.filter((location) => !occupiedLocations.has(location.code));
  const occupied = filteredLocations.filter((location) => occupiedLocations.has(location.code));
  return (
    <div className="page-grid">
      <PageTitle eyebrow="Полки / комірки" title="Координатне зберігання замовлень" text="Кожне замовлення має фізичне місце. Пошук працює по коду локації, зоні, стелажу і полиці." />
      <section className="stats-grid">
        <Metric icon={<Archive />} label="Комірок" value={String(locations.length)} hint="підписані фізично" />
        <Metric icon={<CheckCircle2 />} label="Вільно" value={String(free.length)} hint="система може запропонувати" />
        <Metric icon={<PackageCheck />} label="Зайнято" value={String(occupied.length)} hint="не можна покласти вдруге" />
        <Metric icon={<Search />} label="Пошук" value="Код / зона / полиця" hint="REPAIR-A-01 або READY" />
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Пошук і фільтри</h2><span>{filteredLocations.length} локацій</span></div>
        <div className="action-row">
          <input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Пошук по locationCode" />
          <select value={zoneFilter} onChange={(event) => setZoneFilter(event.target.value as 'ALL' | WarehouseZone)}>
            <option value="ALL">Усі зони</option>
            <option value="REPAIR">REPAIR</option>
            <option value="READY">READY</option>
            <option value="STORAGE">STORAGE</option>
          </select>
          <select value={rackFilter} onChange={(event) => setRackFilter(event.target.value)}>
            <option value="ALL">Усі стелажі</option>
            {[...new Set(locations.map((location) => location.rack))].map((rack) => <option key={rack} value={rack}>{rack}</option>)}
          </select>
          <select value={shelfFilter} onChange={(event) => setShelfFilter(event.target.value)}>
            <option value="ALL">Усі полиці</option>
            {[...new Set(locations.map((location) => location.shelf))].map((shelf) => <option key={shelf} value={shelf}>{shelf}</option>)}
          </select>
        </div>
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Візуальна схема</h2><span>колір = зайнятість</span></div>
        <div className="storage-grid">
          {filteredLocations.map((location) => {
            const order = occupiedLocations.get(location.code);
            return <article className={order ? 'storage-cell storage-cell-busy' : 'storage-cell'} key={location.id}><strong>{location.code}</strong><span>{order ? `${order.id} · ${order.locationStatus === 'У інженера' ? 'у інженера' : 'у комірці'}` : 'вільна'}</span></article>;
          })}
        </div>
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Журнал руху замовлень</h2><span>{movementLogs.length} подій</span></div>
        <div className="table movement-table">
          <div className="table-row table-head"><span>Час</span><span>Замовлення</span><span>Звідки</span><span>Куди</span><span>Хто</span></div>
          {movementLogs.map((entry) => (
            <div className="table-row" key={entry.id}>
              <span>{entry.timestamp}</span>
              <span>{entry.orderId}</span>
              <span>{entry.fromLocation ?? 'Прийом'}</span>
              <span>{entry.toLocation}</span>
              <span>{entry.userId}</span>
            </div>
          ))}
          {movementLogs.length === 0 && <div className="empty-state">Переміщень ще не було.</div>}
        </div>
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Правила розміщення</h2><span>без втрат</span></div>
        <div className="task-list">
          <Task icon={<X />} title="Не можна в зайняту" text="CRM блокує розміщення в комірку, де вже лежить інша одиниця." />
          <Task icon={<Archive />} title="Одна ячейка на цикл" text="Замовлення зберігає одну й ту саму ячейку від прийому до видачі. Змінюється тільки стан: у комірці або у інженера." />
          <Task icon={<Search />} title="Швидкий пошук" text="Пошук працює по QR, замовленню, locationCode, зоні та полиці." />
        </div>
      </section>
    </div>
  );
}

function MovementsPage({ movements, productName }: { movements: StockMovement[]; productName: (id: string) => string }) {
  return (
    <div className="page-grid">
      <PageTitle eyebrow="Рух складу" title="Повна історія змін залишків" text="Кожна операція має тип, дату, запчастину, кількість і зв'язок із ремонтом або закупівлею." />
      <section className="panel">
        <div className="table movement-table">
          <div className="table-row table-head"><span>Дата</span><span>Тип</span><span>Запчастина</span><span>К-сть</span><span>Партія</span><span>Ціна</span><span>Документ</span><span>Підстава</span><span>Хто</span><span>Коментар</span></div>
          {movements.map((movement) => (
            <div className="table-row" key={movement.id}>
              <span>{movement.date}</span>
              <span>{movement.type}</span>
              <span>{productName(movement.productId)}</span>
              <span>{movement.qty}</span>
              <span>{movement.batchRefs ?? '—'}</span>
              <span>{typeof movement.unitPrice === 'number' ? money(movement.unitPrice) : '—'}</span>
              <span>{movement.orderId ?? movement.purchaseId}</span>
              <span>{movement.basis ?? '—'}</span>
              <span>{movement.actor ?? '—'}</span>
              <span>{movement.comment}</span>
            </div>
          ))}
        </div>
        {movements.length === 0 && <div className="empty-state">Рухів ще немає.</div>}
      </section>
    </div>
  );
}

function ClientsPage({
  clients,
  orders,
  simplePayments,
  contracts,
  contractActs,
  onImportClients,
  createContractFromOrders,
  initialSearch,
  focusedClientPhone,
  onFocusedClientPhoneChange,
}: {
  clients: ClientRecord[];
  orders: ServiceOrder[];
  simplePayments: SimpleOrderPaymentRecord[];
  contracts: ContractRecord[];
  contractActs: ContractActRecord[];
  onImportClients: (file: File) => Promise<void>;
  createContractFromOrders: (payload: { client: string; startDate: string; endDate: string; orderIds: string[] }) => boolean;
  initialSearch?: string;
  focusedClientPhone?: string;
  onFocusedClientPhoneChange?: (phone: string) => void;
}) {
  const importClientsInputRef = useRef<HTMLInputElement | null>(null);
  const [clientSearch, setClientSearch] = useState('');
  const [selectedClientPhone, setSelectedClientPhone] = useState('');
  const [selectedOrderIds, setSelectedOrderIds] = useState<string[]>([]);
  const [contractStartDate, setContractStartDate] = useState('');
  const [contractEndDate, setContractEndDate] = useState('');
  const focusClient = (phone: string) => {
    setSelectedClientPhone(phone);
    onFocusedClientPhoneChange?.(phone);
  };

  useEffect(() => {
    if (initialSearch !== undefined) setClientSearch(initialSearch);
  }, [initialSearch]);

  useEffect(() => {
    if (focusedClientPhone !== undefined) setSelectedClientPhone(focusedClientPhone);
  }, [focusedClientPhone]);

  async function handleClientsFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      await onImportClients(file);
    } catch (error) {
      window.alert(error instanceof Error ? error.message : 'Не вдалося імпортувати клієнтів.');
    } finally {
      event.target.value = '';
    }
  }

  const filteredClients = clients.filter((client) => {
    return matchesClientSearch(client, clientSearch);
  });
  const selectedClient = filteredClients.find((client) => client.phone === selectedClientPhone) ?? filteredClients[0] ?? clients[0];

  return (
    <div className="page-grid">
      <PageTitle eyebrow="Клієнти" title="Клієнтська база сервісу" text="Пошук -> список -> вибір -> дія." />
      <section className="panel manager-orders-toolbar">
        <div className="manager-orders-toolbar-search">
          <input value={clientSearch} onChange={(event) => setClientSearch(event.target.value)} placeholder="Пошук: телефон / ім'я / компанія / ЄДРПОУ / ІПН" />
        </div>
        <div className="manager-orders-toolbar-filter">
        <input ref={importClientsInputRef} type="file" accept=".xlsx,.csv" onChange={handleClientsFileChange} style={{ display: 'none' }} />
        <button type="button" onClick={() => importClientsInputRef.current?.click()}>Імпорт клієнтів</button>
        </div>
      </section>
      <div className="manager-orders-workspace">
        <section className="panel manager-orders-list">
          <div className="panel-heading">
            <h2>Клієнти</h2>
            <span>{filteredClients.length}</span>
          </div>
          <div className="manager-orders-list-body">
            <div className="table clients-compact-table">
              <div className="table-row table-head">
                <span>Клієнт</span>
                <span>Телефон</span>
                <span>Замовлення</span>
                <span>Борг</span>
                <span>Останній заказ</span>
                <span>Дія</span>
              </div>
              {filteredClients.map((client, index) => {
                const clientOrders = orders.filter((order) => isSameClientOrder(client, order));
                const lastOrder = [...clientOrders].sort((a, b) => (parseDateTime(b.intakeDate)?.getTime() ?? 0) - (parseDateTime(a.intakeDate)?.getTime() ?? 0))[0];
                const { totalDebt: clientDebt } = clientDebtSummary(client, orders);
                return (
                  <div
                    key={`${client.phone}-${client.email ?? index}`}
                    className={`table-row${selectedClient?.phone === client.phone ? ' is-selected' : ''}`}
                    onClick={() => focusClient(client.phone)}
                    role="button"
                    tabIndex={0}
                    onKeyDown={(event) => {
                      if (event.key === 'Enter' || event.key === ' ') {
                        event.preventDefault();
                        focusClient(client.phone);
                      }
                    }}
                  >
                    <span>{client.name}<small>{client.taxId ? `ЄДРПОУ / ІПН ${client.taxId}` : 'без коду'}</small></span>
                    <span>{client.phone}</span>
                    <span>{client.orders}</span>
                    <span>{money(clientDebt)}</span>
                    <span>{lastOrder ? `${lastOrder.id} · ${lastOrder.intakeDate}` : '—'}</span>
                    <span><button type="button" onClick={(event) => { event.stopPropagation(); focusClient(client.phone); }}>Відкрити</button></span>
                  </div>
                );
              })}
            </div>
            {filteredClients.length === 0 && <div className="empty-state">Клієнтів не знайдено.</div>}
          </div>
        </section>

        <section className="panel manager-order-focus">
          {!selectedClient && <div className="empty-state">Оберіть клієнта зі списку.</div>}
          {selectedClient && (() => {
          const client = selectedClient;
          const uncontractedOrders = clientUncontractedOrders(client, orders);
          const uncontractedAmount = uncontractedOrders.reduce((sum, order) => sum + contractOrderAmount(order), 0);
          const { totalDebt: clientDebt, debtOrdersCount } = clientDebtSummary(client, orders);
          const normalizedTaxId = normalizeTaxId(client.taxId);
          const incomingConfirmed = simplePayments
            .filter((payment) => payment.status === 'подтвержден')
            .filter((payment) => {
              const paymentTaxId = normalizeTaxId(payment.clientTaxId ?? payment.taxId);
              return (normalizedTaxId && paymentTaxId === normalizedTaxId) || payment.client.trim().toLowerCase() === client.name.trim().toLowerCase();
            })
            .reduce((sum, payment) => sum + payment.amount, 0);
          const closedActsTotal = contractActs
            .filter((act) => {
              const contract = contracts.find((item) => item.id === act.contractId);
              return contract?.client.trim().toLowerCase() === client.name.trim().toLowerCase();
            })
            .reduce((sum, act) => sum + act.amount, 0);
          const clientBalance = incomingConfirmed - closedActsTotal;
          const clientOrders = orders.filter((order) => isSameClientOrder(client, order)).sort((a, b) => (parseDateTime(b.intakeDate)?.getTime() ?? 0) - (parseDateTime(a.intakeDate)?.getTime() ?? 0));
          const selectedOrders = uncontractedOrders.filter((order) => selectedOrderIds.includes(order.id));
          const selectedAmount = selectedOrders.reduce((sum, order) => sum + contractOrderAmount(order), 0);
          return (
            <article className="client-card">
              <div className="panel-heading">
                <h2>{client.name}</h2>
                <span>{client.phone}</span>
              </div>
              <div className="manager-order-bottom-meta">
                <div className="manager-order-field"><strong>ЄДРПОУ / ІПН</strong><div>{client.taxId || 'не вказано'}</div></div>
                <div className="manager-order-field"><strong>Email</strong><div>{client.email?.trim() || 'немає email'}</div></div>
                <div className="manager-order-field"><strong>Замовлення</strong><div>{client.orders}</div></div>
                <div className="manager-order-field"><strong>Борг</strong><div>{money(clientDebt)}</div></div>
                <div className="manager-order-field"><strong>Не оформлено</strong><div>{money(uncontractedAmount)}</div></div>
                <div className="manager-order-field"><strong>Баланс</strong><div>{money(clientBalance)}</div></div>
              </div>

              <details className="manager-order-collapse" open>
                <summary>Замовлення клієнта</summary>
                <div className="table clients-compact-table">
                  <div className="table-row table-head">
                    <span>Замовлення</span>
                    <span>Пристрій</span>
                    <span>Статус</span>
                    <span>Сума</span>
                    <span>Борг</span>
                  </div>
                  {clientOrders.map((order) => (
                    <div className="table-row" key={order.id}>
                      <span>{order.id}<small>{order.intakeDate}</small></span>
                      <span>{order.device}</span>
                      <span>{managerOrderStatusLabel(order.status)}</span>
                      <span>{money(contractOrderAmount(order))}</span>
                      <span>{money(clientOrderDebt(order))}</span>
                    </div>
                  ))}
                </div>
              </details>

              <details className="manager-order-collapse">
                <summary>Договори / борги / платежі</summary>
                <div className="manager-order-bottom-meta">
                  <div className="manager-order-field"><strong>Замовлень з боргом</strong><div>{debtOrdersCount}</div></div>
                  <div className="manager-order-field"><strong>Вхідні платежі</strong><div>{money(incomingConfirmed)}</div></div>
                  <div className="manager-order-field"><strong>Закрито актами</strong><div>{money(closedActsTotal)}</div></div>
                </div>
              </details>

              {uncontractedOrders.length > 0 && (
                <details className="manager-order-collapse">
                  <summary>Створити договір із замовлень</summary>
                  <div className="contracts-act-builder">
                    <div className="contracts-act-builder-list">
                      {uncontractedOrders.map((order) => (
                        <label key={order.id} className="contracts-act-order">
                          <input
                            type="checkbox"
                            checked={selectedOrderIds.includes(order.id)}
                            onChange={(event) => {
                              setSelectedOrderIds((current) => (
                                event.target.checked
                                  ? [...current, order.id]
                                  : current.filter((id) => id !== order.id)
                              ));
                            }}
                          />
                          <span>{order.id}</span>
                          <span>{order.device}</span>
                          <strong>{money(contractOrderAmount(order))}</strong>
                        </label>
                      ))}
                    </div>
                    <div className="table">
                      <label>
                        Дата початку
                        <input value={contractStartDate} onChange={(event) => setContractStartDate(event.target.value)} type="date" />
                      </label>
                      <label>
                        Дата завершення
                        <input value={contractEndDate} onChange={(event) => setContractEndDate(event.target.value)} type="date" />
                      </label>
                    </div>
                    <div className="client-meta">
                      <span>Сума договору</span>
                      <strong>{money(selectedAmount)}</strong>
                    </div>
                    <div className="action-row">
                      <button
                        type="button"
                        className="submit-button"
                        onClick={() => {
                          const created = createContractFromOrders({
                            client: client.name,
                            startDate: contractStartDate,
                            endDate: contractEndDate,
                            orderIds: selectedOrderIds,
                          });
                          if (created) {
                            setSelectedOrderIds([]);
                            setContractStartDate('');
                            setContractEndDate('');
                          }
                        }}
                        disabled={selectedOrderIds.length === 0}
                      >
                        Створити договір
                      </button>
                    </div>
                  </div>
                </details>
              )}
            </article>
          );
          })()}
        </section>
      </div>
    </div>
  );
}

function FinancePage({
  orders,
  products,
  receipts,
  movements,
  contracts,
  contractActs,
  bankImportItems,
  simplePayments,
  users,
  onOpenClients,
}: {
  orders: ServiceOrder[];
  products: Product[];
  receipts: GoodsReceipt[];
  movements: StockMovement[];
  contracts: ContractRecord[];
  contractActs: ContractActRecord[];
  bankImportItems: BankImportItem[];
  simplePayments: SimpleOrderPaymentRecord[];
  users: User[];
  onOpenClients: () => void;
}) {
  const [period, setPeriod] = useState<'today' | 'week' | 'month' | 'custom'>('month');
  const [customFrom, setCustomFrom] = useState('');
  const [customTo, setCustomTo] = useState('');
  const [managerFilter, setManagerFilter] = useState('all');
  const [clientTypeFilter, setClientTypeFilter] = useState<'all' | 'physical' | 'legal'>('all');

  const isWithinSelectedPeriod = (dateText: string) => {
    if (period !== 'custom') return isDateInPeriod(dateText, period as AccountingPeriod);
    const parsed = parseDateTime(dateText);
    if (!parsed) return false;
    const from = customFrom ? parseDateTime(customFrom) : null;
    const to = customTo ? parseDateTime(customTo) : null;
    if (from && parsed < from) return false;
    if (to) {
      const inclusiveTo = new Date(to);
      inclusiveTo.setHours(23, 59, 59, 999);
      if (parsed > inclusiveTo) return false;
    }
    return true;
  };

  const orderClientType = (order: ServiceOrder) => (order.legalEntity || order.contractId ? 'legal' : 'physical');
  const managerNameByPayment = (payment: SimpleOrderPaymentRecord) => {
    if (payment.orderId) return orders.find((order) => order.id === payment.orderId)?.manager ?? '';
    if (payment.actId) {
      const act = contractActs.find((item) => item.id === payment.actId);
      const actOrders = orders.filter((order) => act?.orderIds.includes(order.id));
      return actOrders[0]?.manager ?? '';
    }
    return '';
  };
  const clientTypeByPayment = (payment: SimpleOrderPaymentRecord) => {
    if (payment.orderId) return orderClientType(orders.find((order) => order.id === payment.orderId) ?? {} as ServiceOrder);
    if (payment.actId || payment.entityType === 'act' || payment.entityType === 'contract') return 'legal';
    return 'physical';
  };
  const managerMatches = (order: ServiceOrder) => managerFilter === 'all' || order.manager === managerFilter;
  const clientTypeMatches = (order: ServiceOrder) => clientTypeFilter === 'all' || orderClientType(order) === clientTypeFilter;

  const filteredOrders = orders.filter((order) => isWithinSelectedPeriod(order.intakeDate)).filter(managerMatches).filter(clientTypeMatches);
  const visibleOrderIds = new Set(filteredOrders.map((order) => order.id));
  const visibleContracts = contracts.filter((contract) => {
    const linkedOrders = orders.filter((order) => order.contractId === contract.id);
    if (linkedOrders.length === 0) return clientTypeFilter !== 'physical' && managerFilter === 'all' && isWithinSelectedPeriod(contract.createdAt);
    return linkedOrders.some((order) => managerMatches(order) && clientTypeMatches(order) && isWithinSelectedPeriod(order.intakeDate));
  });
  const visibleContractIds = new Set(visibleContracts.map((contract) => contract.id));
  const visibleActs = contractActs.filter((act) => visibleContractIds.has(act.contractId));
  const visibleActIds = new Set(visibleActs.map((act) => act.id));
  const visiblePayments = simplePayments
    .filter((payment) => payment.status === 'подтвержден')
    .filter((payment) => payment.direction !== 'outgoing')
    .filter((payment) => isWithinSelectedPeriod(payment.date))
    .filter((payment) => managerFilter === 'all' || managerNameByPayment(payment) === managerFilter)
    .filter((payment) => clientTypeFilter === 'all' || clientTypeByPayment(payment) === clientTypeFilter);

  const totalOrdersCount = filteredOrders.length;
  const totalOrdersAmount = filteredOrders.reduce((sum, order) => sum + Math.max(order.repairPrice ?? order.estimatedAmount ?? orderTotals(order).total, 0), 0);
  const totalPaid = filteredOrders.reduce((sum, order) => sum + orderDebtSnapshot(order).paid, 0);
  const totalDebt = filteredOrders.reduce((sum, order) => sum + orderDebtSnapshot(order).remainingDebt, 0);
  const unpaidActsAmount = visibleActs.reduce((sum, act) => sum + act.remainingAmount, 0);
  const openContractAmount = visibleContracts.reduce((sum, contract) => sum + Math.max(contractUsage(contract, orders).used - contractClosedAmount(contract, contractActs), 0), 0);
  const totalProfit = filteredOrders.reduce((sum, order) => sum + orderTotals(order).finalProfit, 0);

  const debtorRows = Array.from(
    filteredOrders
      .filter((order) => clientOrderDebt(order) > 0)
      .reduce((map, order) => {
        const key = `${order.client}::${order.phone}`;
        const current = map.get(key) ?? { client: order.client, phone: order.phone, debt: 0, maxDays: 0 };
        current.debt += clientOrderDebt(order);
        current.maxDays = Math.max(current.maxDays, daysSince(order.debtSince ?? order.issuedDebtAt ?? order.statusChangedAt));
        map.set(key, current);
        return map;
      }, new Map<string, { client: string; phone: string; debt: number; maxDays: number }>()),
  )
    .map(([, value]) => value)
    .sort((a, b) => b.debt - a.debt || b.maxDays - a.maxDays)
    .slice(0, 10);

  const actAgeBuckets = visibleActs.reduce((acc, act) => {
    if (act.remainingAmount <= 0) return acc;
    const age = daysSince(act.date);
    if (age <= 7) acc.upto7 += 1;
    else if (age <= 30) acc.upto30 += 1;
    else acc.over30 += 1;
    return acc;
  }, { upto7: 0, upto30: 0, over30: 0 });

  const contractMetrics = visibleContracts.reduce((acc, contract) => {
    const usage = contractUsage(contract, orders);
    const closed = contractClosedAmount(contract, contractActs);
    acc.limit += contract.amount;
    acc.used += usage.used;
    acc.remaining += usage.remaining;
    acc.notClosedByActs += Math.max(usage.used - closed, 0);
    return acc;
  }, { limit: 0, used: 0, remaining: 0, notClosedByActs: 0 });

  const paymentMetrics = {
    amount: visiblePayments.reduce((sum, payment) => sum + payment.amount, 0),
    count: visiblePayments.length,
    unmatched: bankImportItems.filter((item) => item.status === 'unmatched').length,
    review: bankImportItems.filter((item) => item.status === 'review').length,
  };

  const averageProfit = totalOrdersCount > 0 ? totalProfit / totalOrdersCount : 0;
  const topProfitOrders = [...filteredOrders]
    .map((order) => ({ id: order.id, client: order.client, profit: orderTotals(order).finalProfit }))
    .sort((a, b) => b.profit - a.profit)
    .slice(0, 5);

  const stockValue = products.reduce((sum, product) => sum + product.batches.reduce((inner, batch) => inner + batch.qtyAvailable * batch.purchasePrice, 0), 0);
  const staleProducts = products
    .map((product) => {
      const latestMovement = movements
        .filter((movement) => movement.productId === product.id)
        .sort((a, b) => (parseDateTime(b.date)?.getTime() ?? 0) - (parseDateTime(a.date)?.getTime() ?? 0))[0];
      const latestDate = latestMovement?.date ?? product.batches.sort((a, b) => (parseDateTime(b.purchaseDate)?.getTime() ?? 0) - (parseDateTime(a.purchaseDate)?.getTime() ?? 0))[0]?.purchaseDate;
      return { product, days: latestDate ? daysSince(latestDate) : 0 };
    })
    .filter((item) => available(item.product) > 0 && item.days >= 30)
    .sort((a, b) => b.days - a.days);
  const topProductsBySales = Array.from(
    movements
      .filter((movement) => ['Продаж', 'Встановлення'].includes(movement.type))
      .reduce((map, movement) => {
        map.set(movement.productId, (map.get(movement.productId) ?? 0) + movement.qty);
        return map;
      }, new Map<string, number>()),
  )
    .map(([productId, qty]) => ({ productId, qty }))
    .sort((a, b) => b.qty - a.qty)
    .slice(0, 5);

  const redSignals = [
    totalDebt > 0 ? `Долги клієнтів: ${money(totalDebt)}` : '',
    unpaidActsAmount > 0 ? `Неоплачені акти: ${money(unpaidActsAmount)}` : '',
    paymentMetrics.unmatched > 0 ? `Нерозпізнані платежі: ${paymentMetrics.unmatched}` : '',
  ].filter(Boolean);
  const yellowSignals = [
    filteredOrders.filter((order) => {
      const debt = orderDebtSnapshot(order);
      return debt.paid > 0 && debt.remainingDebt > 0;
    }).length > 0 ? `Часткові оплати: ${filteredOrders.filter((order) => {
      const debt = orderDebtSnapshot(order);
      return debt.paid > 0 && debt.remainingDebt > 0;
    }).length}` : '',
    filteredOrders.flatMap((order) => order.payments).filter((payment) => paymentNeedsConfirmation(payment)).length > 0
      ? `Очікують підтвердження: ${filteredOrders.flatMap((order) => order.payments).filter((payment) => paymentNeedsConfirmation(payment)).length}`
      : '',
  ].filter(Boolean);

  const managers = users.filter((user) => user.role === 'Менеджер');

  return (
    <div className="page-grid executive-dashboard finance-center-page">
      <PageTitle eyebrow="Фінанси" title="Фінансовий центр" text="Гроші, борги, акти, договори і склад в одному управлінському екрані." />

      <section className="panel executive-panel finance-center-filters">
        <div className="quick-actions">
          <button type="button" className={period === 'today' ? 'primary' : ''} onClick={() => setPeriod('today')}>Сьогодні</button>
          <button type="button" className={period === 'week' ? 'primary' : ''} onClick={() => setPeriod('week')}>Тиждень</button>
          <button type="button" className={period === 'month' ? 'primary' : ''} onClick={() => setPeriod('month')}>Місяць</button>
          <button type="button" className={period === 'custom' ? 'primary' : ''} onClick={() => setPeriod('custom')}>Період</button>
        </div>
        <div className="finance-center-filter-grid">
          {period === 'custom' && (
            <>
              <label>
                Від
                <input type="date" value={customFrom} onChange={(event) => setCustomFrom(event.target.value)} />
              </label>
              <label>
                До
                <input type="date" value={customTo} onChange={(event) => setCustomTo(event.target.value)} />
              </label>
            </>
          )}
          <label>
            Менеджер
            <select value={managerFilter} onChange={(event) => setManagerFilter(event.target.value)}>
              <option value="all">Усі</option>
              {managers.map((manager) => <option key={manager.id} value={manager.name}>{manager.name}</option>)}
            </select>
          </label>
          <label>
            Тип клієнта
            <select value={clientTypeFilter} onChange={(event) => setClientTypeFilter(event.target.value as 'all' | 'physical' | 'legal')}>
              <option value="all">Усі</option>
              <option value="physical">Фізособи</option>
              <option value="legal">Юрособи</option>
            </select>
          </label>
        </div>
      </section>

      <section className="executive-kpis finance-center-kpis">
        <article className="executive-kpi-card"><span>Всього замовлень</span><strong>{String(totalOrdersCount)}</strong></article>
        <article className="executive-kpi-card"><span>Загальна сума замовлень</span><strong>{money(totalOrdersAmount)}</strong></article>
        <article className="executive-kpi-card"><span>Оплачено</span><strong>{money(totalPaid)}</strong></article>
        <article className="executive-kpi-card"><span>В боргу</span><strong>{money(totalDebt)}</strong></article>
        <article className="executive-kpi-card"><span>В актах, не оплачено</span><strong>{money(unpaidActsAmount)}</strong></article>
        <article className="executive-kpi-card"><span>В договорах, не закрито</span><strong>{money(openContractAmount)}</strong></article>
        <article className="executive-kpi-card"><span>Прибуток</span><strong>{money(totalProfit)}</strong></article>
      </section>

      <section className="manager-orders-signals">
        {redSignals.map((signal) => <span key={signal} className="manager-order-hint manager-order-hint-danger">{signal}</span>)}
        {yellowSignals.map((signal) => <span key={signal} className="manager-order-hint manager-order-hint-warning">{signal}</span>)}
      </section>

      <section className="executive-grid finance-center-top-grid">
        <section className="panel executive-panel">
          <div className="panel-heading">
            <h2>Долги</h2>
            <span>{debtorRows.length} клієнтів</span>
          </div>
          <div className="details-grid">
            <Info label="Загальний борг" value={money(debtorRows.reduce((sum, row) => sum + row.debt, 0))} />
            <Info label="Клієнтів з боргом" value={String(debtorRows.length)} />
          </div>
          <div className="executive-list">
            {debtorRows.map((row) => (
              <button key={`${row.client}-${row.phone}`} type="button" className="executive-list-row finance-center-row-button" onClick={onOpenClients}>
                <span>{row.client}</span>
                <strong>{money(row.debt)} · {row.maxDays} дн.</strong>
              </button>
            ))}
            {debtorRows.length === 0 && <div className="empty-state">Боргів за фільтром немає.</div>}
          </div>
        </section>

        <section className="panel executive-panel">
          <div className="panel-heading">
            <h2>Акти</h2>
            <span>{visibleActs.length} актів</span>
          </div>
          <div className="details-grid">
            <Info label="Сума по актах" value={money(unpaidActsAmount)} />
            <Info label="Кількість" value={String(visibleActs.filter((act) => act.remainingAmount > 0).length)} />
            <Info label="До 7 днів" value={String(actAgeBuckets.upto7)} />
            <Info label="7–30 днів" value={String(actAgeBuckets.upto30)} />
            <Info label="30+ днів" value={String(actAgeBuckets.over30)} />
          </div>
        </section>
      </section>

      <section className="executive-grid finance-center-middle-grid">
        <section className="panel executive-panel">
          <div className="panel-heading">
            <h2>Договори</h2>
            <span>{visibleContracts.length} активних/видимих</span>
          </div>
          <div className="details-grid">
            <Info label="Ліміти" value={money(contractMetrics.limit)} />
            <Info label="Використано" value={money(contractMetrics.used)} />
            <Info label="Залишок" value={money(contractMetrics.remaining)} />
            <Info label="Не закрито актами" value={money(contractMetrics.notClosedByActs)} />
          </div>
        </section>

        <section className="panel executive-panel">
          <div className="panel-heading">
            <h2>Платежі</h2>
            <span>{paymentMetrics.count} за період</span>
          </div>
          <div className="details-grid">
            <Info label="Прийшло грошей" value={money(paymentMetrics.amount)} />
            <Info label="Кількість платежів" value={String(paymentMetrics.count)} />
            <Info label="Нерозпізнані" value={String(paymentMetrics.unmatched)} />
            <Info label="Требують перевірки" value={String(paymentMetrics.review)} />
          </div>
        </section>
      </section>

      <section className="executive-grid finance-center-bottom-grid">
        <section className="panel executive-panel">
          <div className="panel-heading">
            <h2>Прибуток</h2>
            <span>{filteredOrders.length} замовлень</span>
          </div>
          <div className="details-grid">
            <Info label="Загальна прибуток" value={money(totalProfit)} />
            <Info label="Середня з замовлення" value={money(averageProfit)} />
          </div>
          <div className="executive-list">
            {topProfitOrders.map((row) => (
              <div key={row.id} className="executive-list-row">
                <span>{row.id} · {row.client}</span>
                <strong>{money(row.profit)}</strong>
              </div>
            ))}
            {topProfitOrders.length === 0 && <div className="empty-state">Немає прибуткових замовлень за фільтром.</div>}
          </div>
        </section>

        <section className="panel executive-panel">
          <div className="panel-heading">
            <h2>Склад</h2>
            <span>{products.length} позицій</span>
          </div>
          <div className="details-grid">
            <Info label="Залишок товару" value={money(stockValue)} />
            <Info label="Завислий товар 30+" value={String(staleProducts.length)} />
          </div>
          <div className="executive-list">
            {topProductsBySales.map((row) => (
              <div key={row.productId} className="executive-list-row">
                <span>{products.find((product) => product.id === row.productId)?.name ?? row.productId}</span>
                <strong>{row.qty} шт.</strong>
              </div>
            ))}
            {topProductsBySales.length === 0 && <div className="empty-state">Продажів/списань по складу поки немає.</div>}
          </div>
        </section>
      </section>
    </div>
  );
}

function ReportsPage({ orders, sales, purchases, receipts, movements, products, actionLogs, cashShift, canDo, exportAccounting }: { orders: ServiceOrder[]; sales: Sale[]; purchases: PurchaseOrder[]; receipts: GoodsReceipt[]; movements: StockMovement[]; products: Product[]; actionLogs: ActionLog[]; cashShift: CashShift; canDo: (permission: Permission) => boolean; exportAccounting: (kind: AccountingExportKind, format: AccountingExportFormat, period: AccountingPeriod) => void }) {
  const [accountingPeriod, setAccountingPeriod] = useState<AccountingPeriod>('today');
  const salesToday = sales.filter((sale) => sale.date === today);
  const salesRevenue = sales.reduce((sum, sale) => sum + saleTotals(sale).paid, 0);
  const card = sales.flatMap((sale) => sale.payments).filter((payment) => payment.method === 'Картка').reduce((sum, payment) => sum + payment.amount, 0);
  const cash = sales.flatMap((sale) => sale.payments).filter((payment) => payment.method === 'Готівка').reduce((sum, payment) => sum + payment.amount, 0);
  const bank = sales.flatMap((sale) => sale.payments).filter((payment) => payment.method === 'Безготівка').reduce((sum, payment) => sum + payment.amount, 0);
  const managerRevenue = sales.reduce<Record<string, number>>((acc, sale) => {
    acc[sale.manager] = (acc[sale.manager] ?? 0) + saleTotals(sale).paid;
    return acc;
  }, {});
  const returns = sales.flatMap((sale) => sale.returns);
  const productProfit = sales.reduce((sum, sale) => sum + saleTotals(sale).profit, 0);
  const repairProfit = orders.reduce((sum, order) => sum + orderTotals(order).finalProfit, 0);
  const closedOrders = orders.filter((order) => ['Закрито', 'Видано'].includes(order.status)).length;
  const unpaidOrders = orders.filter((order) => orderTotals(order).debt > 0).length;
  const stuckOrders = orders.filter((order) => daysSince(order.statusChangedAt) > 2 && !['Закрито', 'Видано', 'Скасовано'].includes(order.status)).length;
  const profitByDay = orders.reduce<Record<string, number>>((acc, order) => {
    const key = extractDayKey(order.statusChangedAt);
    acc[key] = (acc[key] ?? 0) + orderTotals(order).finalProfit;
    return acc;
  }, {});
  const profitByMonth = orders.reduce<Record<string, number>>((acc, order) => {
    const key = extractMonthKey(order.statusChangedAt);
    acc[key] = (acc[key] ?? 0) + orderTotals(order).finalProfit;
    return acc;
  }, {});
  const accountingPayments = [
    ...orders.flatMap((order) => order.payments.filter((payment) => paymentCountsAsApplied(payment) && isDateInPeriod(payment.date, accountingPeriod))),
    ...sales.flatMap((sale) => sale.payments.filter((payment) => paymentCountsAsApplied(payment) && isDateInPeriod(payment.date, accountingPeriod))),
  ];
  const accountingTurnover = accountingPayments.reduce((sum, payment) => sum + payment.amount, 0);
  const accountingPurchases = receipts.filter((receipt) => isDateInPeriod(receipt.date, accountingPeriod)).reduce((sum, receipt) => sum + receipt.qty * receipt.price, 0);
  const accountingDebt = orders.reduce((sum, order) => sum + orderTotals(order).debt, 0) + sales.reduce((sum, sale) => sum + saleTotals(sale).debt, 0);
  const accountingStockMoves = movements.filter((movement) => isDateInPeriod(movement.date, accountingPeriod)).length;
  const visiblePurchases = purchases.filter((purchase) => isDateInPeriod(purchase.orderedAt, accountingPeriod)).length;

  return (
    <div className="page-grid">
      <PageTitle eyebrow="Звіти" title="Продажі, каса, борги, повернення і прибуток" text="Звіти рахуються автоматично з оплат, видач, повернень, рухів складу та закритих ремонтів." />
      <section className="stats-grid">
        <Metric icon={<ShoppingCart />} label="Продажі за день" value={String(salesToday.length)} hint={money(salesRevenue)} />
        <Metric icon={<Banknote />} label="Готівка" value={money(cash)} hint="каса" />
        <Metric icon={<CheckCircle2 />} label="Картка" value={money(card)} hint="термінал" />
        <Metric icon={<Archive />} label="Безготівка" value={money(bank)} hint="рахунок" />
      </section>
      <section className="stats-grid">
        <Metric icon={<PackageCheck />} label="Закрито замовлень" value={String(closedOrders)} hint="реальні завершені події" />
        <Metric icon={<Banknote />} label="Не оплачено" value={String(unpaidOrders)} hint="замовлення з боргом" />
        <Metric icon={<History />} label="Зависло" value={String(stuckOrders)} hint="без руху більше 2 днів" />
        <Metric icon={<Wrench />} label="Прибуток ремонтів" value={canDo('reports.profit') || canDo('profit.view') ? money(repairProfit) : 'Приховано'} hint="по закритих замовленнях" />
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Вигрузка в бухгалтерію</h2><span>CRM = джерело даних, BAS / 1С = облік і звітність</span></div>
        <div className="quick-actions">
          <button type="button" className={accountingPeriod === 'today' ? 'primary' : ''} onClick={() => setAccountingPeriod('today')}>Сьогодні</button>
          <button type="button" className={accountingPeriod === 'week' ? 'primary' : ''} onClick={() => setAccountingPeriod('week')}>Тиждень</button>
          <button type="button" className={accountingPeriod === 'month' ? 'primary' : ''} onClick={() => setAccountingPeriod('month')}>Місяць</button>
        </div>
        <div className="details-grid">
          <Info label="Надійшло грошей" value={money(accountingTurnover)} />
          <Info label="Закупки за період" value={money(accountingPurchases)} />
          <Info label="Борги клієнтів" value={money(accountingDebt)} />
          <Info label="Рухів складу" value={String(accountingStockMoves)} />
          <Info label="Закупок у періоді" value={String(visiblePurchases)} />
          <Info label="Формати" value="Excel, CSV, JSON" />
        </div>
        <div className="action-row" style={{ marginTop: '16px', flexWrap: 'wrap' }}>
          <button type="button" disabled={!canDo('reports.export')} onClick={() => exportAccounting('summary', 'excel', accountingPeriod)}>Скачати звіт</button>
          <button type="button" disabled={!canDo('reports.export')} onClick={() => exportAccounting('sales', 'excel', accountingPeriod)}>Продажі Excel</button>
          <button type="button" disabled={!canDo('reports.export')} onClick={() => exportAccounting('sales', 'csv', accountingPeriod)}>Продажі CSV</button>
          <button type="button" disabled={!canDo('reports.export')} onClick={() => exportAccounting('sales', 'json', accountingPeriod)}>Продажі JSON</button>
          <button type="button" disabled={!canDo('reports.export')} onClick={() => exportAccounting('purchases', 'excel', accountingPeriod)}>Закупки Excel</button>
          <button type="button" disabled={!canDo('reports.export')} onClick={() => exportAccounting('purchases', 'csv', accountingPeriod)}>Закупки CSV</button>
          <button type="button" disabled={!canDo('reports.export')} onClick={() => exportAccounting('stock', 'csv', accountingPeriod)}>Рух складу CSV</button>
          <button type="button" disabled={!canDo('reports.export')} onClick={() => exportAccounting('stock', 'json', accountingPeriod)}>Рух складу JSON</button>
        </div>
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Інтеграція з BAS</h2><span>платформа = операційна система</span></div>
        <div className="task-list">
          <Task icon={<ClipboardList />} title="Що передаємо" text="Продажі з розділенням робіт і запчастин, закупки, рух складу і коротке зведення за період." />
          <Task icon={<Archive />} title="Перший етап" text="Бухгалтер натискає одну кнопку і отримує Excel / CSV / JSON без ручного збирання таблиць." />
          <Task icon={<CheckCircle2 />} title="Не робимо в CRM" text="Проводки, рахунки 361/631, податки і фінальна звітність залишаються у BAS / 1С." />
          <Task icon={<Banknote />} title="Роль BAS" text="BAS використовується для бухгалтерії і закону; CRM лишається джерелом чистих операційних даних." />
        </div>
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Каса і термінал</h2><span>{cashShift.id}</span></div>
        <div className="details-grid">
          <Info label="Відкриття зміни" value={cashShift.openedAt} />
          <Info label="Статус" value={cashShift.status} />
          <Info label="Готівка в касі" value={money(expectedCashAmount(cashShift))} />
          <Info label="Термінал" value={money(cashShift.cardIncome)} />
          <Info label="Безготівка" value={money(cashShift.bankIncome)} />
          <Info label="Повернення готівкою" value={money(cashShift.cashExpense)} />
        </div>
      </section>
      <section className="content-split">
        <div className="panel">
          <div className="panel-heading"><h2>Виручка по менеджерах</h2><span>{Object.keys(managerRevenue).length} співробітників</span></div>
          <div className="task-list">
            {Object.entries(managerRevenue).map(([manager, value]) => <Task key={manager} icon={<Users />} title={manager} text={`Виручка: ${money(value)}`} />)}
          </div>
        </div>
        <div className="panel">
          <div className="panel-heading"><h2>Прибуток</h2><span>товари і ремонти</span></div>
          <div className="task-list">
            <Task icon={<PackageCheck />} title="Прибуток по товарах" text={canDo('reports.profit') || canDo('profit.view') ? money(productProfit) : 'Приховано правами доступу'} />
            <Task icon={<Wrench />} title="Прибуток по ремонтах" text={canDo('reports.profit') || canDo('profit.view') ? money(repairProfit) : 'Приховано правами доступу'} />
            <Task icon={<Truck />} title="Повернення" text={`${returns.length} документів`} />
          </div>
        </div>
      </section>
      <section className="content-split">
        <div className="panel">
          <div className="panel-heading"><h2>Прибуток по днях</h2><span>події закриття</span></div>
          <div className="task-list">
            {Object.entries(profitByDay).slice(0, 6).map(([day, amount]) => <Task key={day} icon={<Banknote />} title={day} text={canDo('reports.profit') || canDo('profit.view') ? money(amount) : 'Приховано правами доступу'} />)}
          </div>
        </div>
        <div className="panel">
          <div className="panel-heading"><h2>Прибуток по місяцях</h2><span>агреговано</span></div>
          <div className="task-list">
            {Object.entries(profitByMonth).slice(0, 6).map(([month, amount]) => <Task key={month} icon={<Archive />} title={month} text={canDo('reports.profit') || canDo('profit.view') ? money(amount) : 'Приховано правами доступу'} />)}
          </div>
        </div>
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Контроль складу і логів</h2><span>{movements.length} рухів · {actionLogs.length} дій</span></div>
        <div className="table movement-table">
          <div className="table-row table-head"><span>Дата</span><span>Користувач</span><span>Роль</span><span>Дія</span><span>Документ</span><span>Коментар</span></div>
          {actionLogs.slice(0, 8).map((log) => (
            <div className="table-row" key={log.id}>
              <span>{log.date}</span>
              <span>{log.user}</span>
              <span>{log.role}</span>
              <span>{log.action}</span>
              <span>{log.entity}</span>
              <span>{log.comment}</span>
            </div>
          ))}
        </div>
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Залишок після продажів</h2><span>{products.length} позицій</span></div>
        <div className="task-list">
          {products.map((product) => <Task key={product.id} icon={<Archive />} title={product.name} text={`Доступно ${available(product)} шт., резерв ${product.reserved} шт.`} />)}
        </div>
      </section>
    </div>
  );
}

function PayrollPage({ orders, activeUser }: { orders: ServiceOrder[]; activeUser: User }) {
  const closedStatuses: OrderStatus[] = ['Видано', 'Закрито'];
  const [periodFilter, setPeriodFilter] = useState<'today' | 'week' | 'month'>('today');
  const todayDate = parseDateTime(today) ?? new Date();
  const isInPeriod = (dateText?: string) => {
    if (!dateText) return false;
    const date = parseDateTime(dateText);
    if (!date) return false;
    const diffDays = Math.floor((todayDate.getTime() - date.getTime()) / 86400000);
    if (periodFilter === 'today') return extractDayKey(dateText) === extractDayKey(today);
    if (periodFilter === 'week') return diffDays >= 0 && diffDays < 7;
    return extractMonthKey(dateText) === extractMonthKey(today);
  };
  const workRows = orders.flatMap((order) =>
    order.works.map((work, index) => {
      const employee = work.engineer ?? order.engineer;
      const rule = payrollRules.find((item) => item.employee === employee);
      const qty = work.qty ?? 1;
      const workAmount = orderWorkAmount(work);
      const paidAndClosed = closedStatuses.includes(order.status) && orderTotals(order).debt <= 0 && orderTotals(order).paid > 0;
      const accrued = paidAndClosed ? payrollForWork(order, work, rule) : 0;
      const completedAt = paidAndClosed ? order.statusChangedAt : undefined;
      return {
        id: `${order.id}-${index}`,
        order,
        work,
        employee,
        rule,
        workType: workTypeLabel(work),
        qty,
        workAmount,
        rateLabel: workRateLabel(work, rule),
        accrued,
        completedAt,
        status: paidAndClosed ? 'Нараховано' : 'Очікує оплату / видачу',
      };
    }),
  );
  const visibleRows = activeUser.role === 'Інженер' ? workRows.filter((row) => row.employee === activeUser.name) : workRows;
  const employeeSummaries = buildPayrollEmployeeSummaries(orders).filter((row) => activeUser.role === 'Інженер' ? row.rule.employee === activeUser.name : true);
  const todayKey = extractDayKey(today);
  const monthKey = extractMonthKey(today);
  const completedVisibleRows = visibleRows.filter((row) => row.status === 'Нараховано');
  const filteredCompletedRows = completedVisibleRows.filter((row) => isInPeriod(row.completedAt));
  const totalAccrued = completedVisibleRows.reduce((sum, row) => sum + row.accrued, 0);
  const totalWorkAmount = completedVisibleRows.reduce((sum, row) => sum + row.workAmount, 0);
  const pendingRows = visibleRows.filter((row) => row.status === 'Очікує оплату / видачу');
  const filteredRepairAccrued = filteredCompletedRows.filter((row) => row.workType === 'Ремонт техники').reduce((sum, row) => sum + row.accrued, 0);
  const filteredCartridgeAccrued = filteredCompletedRows.filter((row) => row.workType !== 'Ремонт техники').reduce((sum, row) => sum + row.accrued, 0);
  const visibleDayAccrued = completedVisibleRows.filter((row) => row.completedAt && extractDayKey(row.completedAt) === todayKey).reduce((sum, row) => sum + row.accrued, 0);
  const visibleMonthAccrued = completedVisibleRows.filter((row) => row.completedAt && extractMonthKey(row.completedAt) === monthKey).reduce((sum, row) => sum + row.accrued, 0);
  const visibleDayOperations = completedVisibleRows.filter((row) => row.completedAt && extractDayKey(row.completedAt) === todayKey).length;
  const visibleMonthOperations = completedVisibleRows.filter((row) => row.completedAt && extractMonthKey(row.completedAt) === monthKey).length;
  const isEngineer = activeUser.role === 'Інженер';
  const isSupervisor = activeUser.role === 'Руководитель' || activeUser.role === 'Адміністратор';

  return (
    <div className="page-grid">
      <PageTitle eyebrow="Зарплата" title={isEngineer ? `Мой заработок: ${activeUser.name}` : 'Начисления сотрудников'} text="Зарплата начисляется автоматически только после оплаты и выдачи заказа: ремонт техники считается процентом только от работы, а заправка и регенерация картриджей считаются по ставке за штуку." />
      <section className="quick-actions">
        <button type="button" onClick={() => setPeriodFilter('today')}>Сьогодні</button>
        <button type="button" onClick={() => setPeriodFilter('week')}>Тиждень</button>
        <button type="button" onClick={() => setPeriodFilter('month')}>Місяць</button>
      </section>
      <section className="stats-grid">
        <Metric icon={<Banknote />} label="Ремонт" value={money(filteredRepairAccrued)} hint="тільки від оплаченої роботи" />
        <Metric icon={<Banknote />} label="Картриджі" value={money(filteredCartridgeAccrued)} hint="заправка / регенерація" />
        <Metric icon={<Wrench />} label="Усього" value={money(filteredCompletedRows.reduce((sum, row) => sum + row.accrued, 0))} hint={`${filteredCompletedRows.length} нарахованих робіт`} />
        <Metric icon={<History />} label="Очікує грошей" value={String(pendingRows.length)} hint="ще не входить у зарплату" />
      </section>
      <section className="content-split">
        <div className="panel">
          <div className="panel-heading"><h2>{isEngineer ? 'Підсумок по мені' : 'Інженери за період'}</h2><span>{employeeSummaries.length} людей</span></div>
          <div className="task-list">
            {employeeSummaries.map((row) => (
              <details key={row.rule.employee} className="task">
                <summary style={{ cursor: 'pointer', fontWeight: 700 }}>{row.rule.employee} — замовлень {row.ordersCount} · ремонт {money(row.repairAccrued)} · картриджі {money(row.cartridgeAccrued)} · всього {money(row.accrued)}</summary>
                <div style={{ marginTop: '8px' }}>
                  <p>Середній дохід на замовлення: {money(row.avgOrderAccrued)}</p>
                  {visibleRows.filter((item) => item.employee === row.rule.employee && item.status === 'Нараховано').map((item) => (
                    <p key={`detail-${item.id}`}>{item.order.id} · {item.workType} · {money(item.workAmount)} → {money(item.accrued)}</p>
                  ))}
                </div>
              </details>
            ))}
          </div>
        </div>
        <div className="panel">
          <div className="panel-heading"><h2>{isEngineer ? 'Як рахується' : 'Правила розрахунку'}</h2><span>тільки після грошей</span></div>
          <div className="task-list">
            {payrollRules.filter((rule) => activeUser.role === 'Інженер' ? rule.employee === activeUser.name : true).map((rule) => (
              <Task key={rule.employee} icon={<ShieldCheck />} title={rule.employee} text={rule.role === 'Інженер' ? `Ремонт техніки: ${rule.percent ?? 0}% тільки від роботи. Заправка і регенерація рахуються по ставці за штуку. Якщо гроші не отримані — зарплати немає.` : `${rule.type}.`} />
            ))}
            {employeeSummaries.length === 0 && <div className="empty-state">Пока нет начислений для отображения.</div>}
          </div>
        </div>
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>{isEngineer ? 'Мої нарахування' : 'Список нарахувань'}</h2><span>{filteredCompletedRows.length} рядків за період</span></div>
        <div className="table payroll-table">
          <div className="table-row table-head"><span>Сотрудник</span><span>Заказ</span><span>Тип работы</span><span>Исполнитель / работа</span><span>Количество</span><span>Ставка</span><span>Сумма работы</span><span>Заработок</span></div>
          {filteredCompletedRows.map((row) => (
            <div className="table-row" key={row.id}>
              <span>{row.employee}</span>
              <span>{row.order.id}<small>{row.status}</small></span>
              <span>{row.workType}</span>
              <span>{row.work.name}<small>{row.completedAt ?? row.order.statusChangedAt}</small></span>
              <span>{row.qty}</span>
              <span>{row.rateLabel}</span>
              <span>{money(row.workAmount)}</span>
              <span>{money(row.accrued)}</span>
            </div>
          ))}
        </div>
      </section>
      {isSupervisor && <section className="panel">
        <div className="panel-heading"><h2>Підсумки для керівника</h2><span>тільки закриті та оплачені замовлення</span></div>
        <div className="task-list">
          {employeeSummaries.map((row) => (
            <Task key={`owner-payroll-${row.rule.employee}`} icon={<Banknote />} title={row.rule.employee} text={`За день ${money(row.dayAccrued)}, за місяць ${money(row.monthAccrued)}, ремонт ${money(row.repairAccrued)}, картриджі ${money(row.cartridgeAccrued)}, середній дохід ${money(row.avgOrderAccrued)}.`} />
          ))}
        </div>
      </section>}
    </div>
  );
}

function AcceptanceChecklistPage() {
  const done = acceptanceChecklist.filter((item) => item.done).length;
  const total = acceptanceChecklist.length;
  return (
    <div className="page-grid">
      <PageTitle eyebrow="Чек-лист приймання" title="Контроль перед демонстрацією" text="Цей список використовується як фінальна перевірка системи перед здачею. Система вважається готовою тільки після проходження всіх пунктів." />
      <section className="stats-grid">
        <Metric icon={<CheckCircle2 />} label="Готово" value={`${done}/${total}`} hint="поточний MVP" />
        <Metric icon={<ShieldCheck />} label="Безпека" value="Логи і ролі" hint="закриті документи захищені" />
        <Metric icon={<Archive />} label="Backup" value="Cloud етап" hint="щоденний backup на VPS" />
        <Metric icon={<Wrench />} label="Анти-1С" value="1-3 кліки" hint="плоске меню" />
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Пункти приймання</h2><span>{total} перевірок</span></div>
        <div className="task-list">
          {acceptanceChecklist.map((item) => (
            <Task key={`${item.group}-${item.item}`} icon={item.done ? <CheckCircle2 /> : <History />} title={`${item.group} · ${item.done ? 'готово' : 'потрібен backend/cloud'}`} text={item.item} />
          ))}
        </div>
      </section>
      <section className="panel">
        <div className="panel-heading"><h2>Архітектура</h2><span>цільовий запуск</span></div>
        <div className="task-list">
          <Task icon={<CheckCircle2 />} title="Веб-доступ" text="CRM має працювати через браузер з офісу, віддалено і з телефона." />
          <Task icon={<Archive />} title="Одна база" text="Усі користувачі працюють в одній PostgreSQL базі на VPS або cloud." />
          <Task icon={<ShieldCheck />} title="Захист" text="HTTPS, логін/пароль, ролі, сесії, журнал дій і щоденне резервне копіювання." />
        </div>
      </section>
    </div>
  );
}

function TeamPage({ users, setUsers, activeUser }: { users: User[]; setUsers: React.Dispatch<React.SetStateAction<User[]>>; activeUser: User }) {
  const [name, setName] = useState('');
  const [phone, setPhone] = useState('');
  const [email, setEmail] = useState('');
  const [role, setRole] = useState<Role>('Інженер');
  const [permissions, setPermissions] = useState<UserPermissions>(() => defaultPermissionsForRole('Інженер'));
  const [search, setSearch] = useState('');
  const [filter, setFilter] = useState<'all' | 'warehouse' | 'manager' | 'engineer'>('all');
  const [editingUserId, setEditingUserId] = useState('');
  const [editName, setEditName] = useState('');
  const [editPhone, setEditPhone] = useState('');
  const [editEmail, setEditEmail] = useState('');
  const [editRole, setEditRole] = useState<Role>('Інженер');
  const [editPermissions, setEditPermissions] = useState<UserPermissions>(() => defaultPermissionsForRole('Інженер'));
  const visibleUsers = users.filter((person) => person.login !== 'dev.admin');
  const filteredVisibleUsers = visibleUsers.filter((person) => {
    const needle = search.trim().toLowerCase();
    if (needle) {
      const haystack = [person.name, person.phone ?? '', person.email ?? '', person.role].join(' ').toLowerCase();
      if (!haystack.includes(needle)) return false;
    }
    if (filter === 'warehouse') return person.permissions.canAccessWarehouse;
    if (filter === 'manager') return person.role === 'Менеджер';
    if (filter === 'engineer') return person.role === 'Інженер';
    return true;
  });
  const canManageEmployees = activeUser.role === 'Адміністратор';
  const simplifiedRoles: Array<{ value: Role; label: string }> = baseEmployeeRoles.map((value) => ({
    value,
    label: value === 'Адміністратор' ? 'Адмін' : value,
  }));
  const permissionLabels: Array<{ key: keyof UserPermissions; label: string }> = [
    { key: 'canAccessWarehouse', label: 'Склад' },
    { key: 'canAccessFinance', label: 'Фінанси' },
    { key: 'canAccessEmployees', label: 'Співробітники' },
    { key: 'canAccessSettings', label: 'Налаштування' },
    { key: 'canAccessReports', label: 'Звіти' },
  ];

  const handleRoleChange = (nextRole: Role) => {
    setRole(nextRole);
    setPermissions(defaultPermissionsForRole(nextRole));
  };

  const handleEditRoleChange = (nextRole: Role) => {
    setEditRole(nextRole);
    setEditPermissions(defaultPermissionsForRole(nextRole));
  };

  const startEditing = (person: User) => {
    setEditingUserId(person.id);
    setEditName(person.name);
    setEditPhone(person.phone ?? '');
    setEditEmail(person.email ?? '');
    setEditRole(person.role);
    setEditPermissions({ ...person.permissions });
  };

  const cancelEditing = () => {
    setEditingUserId('');
    setEditName('');
    setEditPhone('');
    setEditEmail('');
    setEditRole('Інженер');
    setEditPermissions(defaultPermissionsForRole('Інженер'));
  };

  const createEmployee = () => {
    const trimmedName = name.trim();
    if (!trimmedName) return;
    const loginBase = (email.trim() || phone.trim() || trimmedName).toLowerCase().replace(/[^a-z0-9а-яіїєґ]+/gi, '.');
    setUsers((current) => {
      const updatedUsers = normalizeUsers([
        {
          id: uid('u'),
          name: trimmedName,
          phone: phone.trim() || undefined,
          email: email.trim() || undefined,
          role,
          twoFactor: false,
          login: loginBase,
          authMode: 'phone_code',
          permissions,
          session: 'Активна',
        },
        ...current,
      ]);
      saveEmployeesToStorage(updatedUsers);
      return updatedUsers;
    });
    setName('');
    setPhone('');
    setEmail('');
    setRole('Інженер');
    setPermissions(defaultPermissionsForRole('Інженер'));
  };

  const saveEmployee = () => {
    const trimmedName = editName.trim();
    if (!trimmedName || !editingUserId) return;
    setUsers((current) => {
      const updatedUsers = normalizeUsers(current.map((person) => (
        person.id === editingUserId
          ? {
              ...person,
              name: trimmedName,
              phone: editPhone.trim() || undefined,
              email: editEmail.trim() || undefined,
              role: editRole,
              permissions: editPermissions,
            }
          : person
      )));
      saveEmployeesToStorage(updatedUsers);
      return updatedUsers;
    });
    cancelEditing();
  };

  const deleteEmployee = (person: User) => {
    if (person.id === activeUser.id) {
      window.alert('Не можна видалити поточного адміністратора із активної сесії.');
      return;
    }
    if (!window.confirm(`Видалити співробітника ${person.name}?`)) return;
    setUsers((current) => {
      const updatedUsers = current.filter((item) => item.id !== person.id);
      saveEmployeesToStorage(updatedUsers);
      return updatedUsers;
    });
    if (editingUserId === person.id) cancelEditing();
  };

  return (
    <div className="page-grid">
      <section className="panel">
        <div className="panel-heading"><h2>Створити співробітника</h2></div>
        <div className="table">
          <label>
            Імʼя
            <input value={name} onChange={(event) => setName(event.target.value)} placeholder="ПІБ співробітника" />
          </label>
          <label>
            Телефон
            <input value={phone} onChange={(event) => setPhone(event.target.value)} placeholder="+380..." />
          </label>
          <label>
            Email
            <input value={email} onChange={(event) => setEmail(event.target.value)} placeholder="необов’язково" />
          </label>
          <label>
            Роль
            <select value={role} onChange={(event) => handleRoleChange(event.target.value as Role)}>
              {simplifiedRoles.map((item) => <option key={item.value} value={item.value}>{item.label}</option>)}
            </select>
          </label>
        </div>
        <div className="task-list" style={{ marginTop: '16px' }}>
          {permissionLabels.map((item) => (
            <label key={item.key} style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
              <input
                type="checkbox"
                checked={permissions[item.key]}
                onChange={(event) => setPermissions((current) => ({ ...current, [item.key]: event.target.checked }))}
              />
              <span>{item.label}</span>
            </label>
          ))}
        </div>
        <div className="action-row">
          <button type="button" className="submit-button" onClick={createEmployee}>Створити співробітника</button>
        </div>
      </section>
      <section className="cards-grid">
        <div className="panel" style={{ gridColumn: '1 / -1' }}>
          <label style={{ display: 'grid', gap: '8px', marginBottom: '12px' }}>
            <input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Пошук по імені, телефону, email або ролі" />
          </label>
          <div className="action-row">
            <button type="button" onClick={() => setFilter('all')}>Усі</button>
            <button type="button" onClick={() => setFilter('warehouse')}>Тільки зі складом</button>
            <button type="button" onClick={() => setFilter('manager')}>Тільки менеджери</button>
            <button type="button" onClick={() => setFilter('engineer')}>Тільки інженери</button>
          </div>
        </div>
        {filteredVisibleUsers.length === 0 && <div className="panel empty-state">Немає співробітників за цим фільтром</div>}
        {filteredVisibleUsers.map((person) => (
          <article className="team-card" key={person.id}>
            <div className="avatar">{person.name.slice(0, 1)}</div>
            {editingUserId === person.id ? (
              <>
                <div className="table">
                  <label>
                    Імʼя
                    <input value={editName} onChange={(event) => setEditName(event.target.value)} />
                  </label>
                  <label>
                    Телефон
                    <input value={editPhone} onChange={(event) => setEditPhone(event.target.value)} />
                  </label>
                  <label>
                    Email
                    <input value={editEmail} onChange={(event) => setEditEmail(event.target.value)} />
                  </label>
                  <label>
                    Роль
                    <select value={editRole} onChange={(event) => handleEditRoleChange(event.target.value as Role)}>
                      {simplifiedRoles.map((item) => <option key={item.value} value={item.value}>{item.label}</option>)}
                    </select>
                  </label>
                </div>
                <div className="task-list" style={{ marginTop: '12px' }}>
                  {permissionLabels.map((item) => (
                    <label key={`edit-${person.id}-${item.key}`} style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
                      <input
                        type="checkbox"
                        checked={editPermissions[item.key]}
                        onChange={(event) => setEditPermissions((current) => ({ ...current, [item.key]: event.target.checked }))}
                      />
                      <span>{item.label}</span>
                    </label>
                  ))}
                </div>
                <div className="action-row" style={{ marginTop: '12px' }}>
                  <button type="button" className="submit-button" onClick={saveEmployee}>Зберегти</button>
                  <button type="button" onClick={cancelEditing}>Скасувати</button>
                </div>
              </>
            ) : (
              <>
                <h2>{person.name}</h2>
                <p>{roleDisplay(person.role)}</p>
                <div className="client-meta">
                  <span>{person.phone ?? 'без телефону'}</span>
                  <span>{person.email?.trim() || 'немає email'}</span>
                  <span>{authModeLabel(person.authMode)}</span>
                  <span>{person.session}</span>
                  <span>Склад: {person.permissions.canAccessWarehouse ? 'так' : 'ні'}</span>
                  <span>Фінанси: {person.permissions.canAccessFinance ? 'так' : 'ні'}</span>
                  <span>Співробітники: {person.permissions.canAccessEmployees ? 'так' : 'ні'}</span>
                  <span>Налаштування: {person.permissions.canAccessSettings ? 'так' : 'ні'}</span>
                  <span>Звіти: {person.permissions.canAccessReports ? 'так' : 'ні'}</span>
                </div>
                {canManageEmployees && (
                  <div className="action-row" style={{ marginTop: '12px' }}>
                    <button type="button" onClick={() => startEditing(person)}>Редагувати</button>
                    <button type="button" onClick={() => deleteEmployee(person)}>Видалити</button>
                  </div>
                )}
              </>
            )}
          </article>
        ))}
      </section>
    </div>
  );
}

function SettingsPage({
  actionLogs,
  activeUser,
  backups,
  createManualBackup,
  restoreBackup,
  exportBackup,
  exportCurrentLiveData,
  importBackupFile,
}: {
  actionLogs: ActionLog[];
  activeUser: User;
  backups: BackupSnapshot[];
  createManualBackup: () => void;
  restoreBackup: (snapshotId: string) => void;
  exportBackup: (snapshotId: string) => void;
  exportCurrentLiveData: () => void;
  importBackupFile: (file: File) => Promise<void>;
}) {
  const backupInputRef = useRef<HTMLInputElement | null>(null);
  return (
    <div className="page-grid">
      <PageTitle eyebrow="Налаштування" title="Правила обліку СПЕКТР-АС" text="Система забороняє резерв без залишку, списання без документа і пряме ручне редагування складу." />
      <section className="settings-grid">
        <div className="panel"><h2>Компанія</h2><dl className="details"><div><dt>Назва</dt><dd>{companySettings.companyName}</dd></div><div><dt>Основний email</dt><dd>{companySettings.mainEmail}</dd></div><div><dt>ЄДРПОУ</dt><dd>{companySettings.edrpou}</dd></div><div><dt>ІПН</dt><dd>{companySettings.ipn}</dd></div><div><dt>IBAN</dt><dd>{companySettings.iban}</dd></div><div><dt>Банк</dt><dd>{companySettings.bank}</dd></div><div><dt>МФО</dt><dd>{companySettings.mfo}</dd></div><div><dt>Адреса</dt><dd>{companySettings.address}</dd></div><div><dt>Телефон</dt><dd>{companySettings.phone}</dd></div><div><dt>ПДВ</dt><dd>{companySettings.vatEnabled ? 'Увімкнено' : 'Вимкнено'}</dd></div><div><dt>Свідоцтво ПДВ</dt><dd>{companySettings.vatCertificate}</dd></div><div><dt>Валюта</dt><dd>Українська гривня</dd></div></dl></div>
        <div className="panel"><h2>Обов'язкові документи</h2><div className="task-list"><Task icon={<PackagePlus />} title="Потреба" text="Створюється, якщо запчастини немає на складі." /><Task icon={<ShoppingCart />} title="Закупівля" text="Пов'язується з ремонтом і потребою." /><Task icon={<History />} title="Рух складу" text="Фіксує кожну зміну залишків." /></div></div>
        <div className="panel"><h2>Анти-1С інтерфейс</h2><div className="task-list"><Task icon={<CheckCircle2 />} title="1-3 кліки" text="Головні дії винесені великими кнопками: прийняти оплату, видати, взяти в роботу, друк квитанції." /><Task icon={<LayoutDashboard />} title="Плоске меню" text="Меню веде одразу на сторінку, без підменю і підподменю." /><Task icon={<Plus />} title="Складне сховано" text="Рідкі дії відкриваються через кнопку Показати ще." /></div></div>
        <div className="panel"><h2>Уведомлення клієнтам</h2><div className="task-list">{clientNotificationTemplates.slice(0, 4).map((template) => <Task key={`${template.event}-${template.channel}`} icon={<Bell />} title={`${template.event} · ${template.channel}`} text={`${template.enabled ? 'Увімкнено' : 'Вимкнено'}: ${template.text}`} />)}</div></div>
      </section>
      <section className="panel">
        <div className="panel-heading">
          <h2>Резервні копії CRM</h2>
          <span>{backups.length} копій</span>
        </div>
        <div className="manager-order-inline-actions">
          <button type="button" className="submit-button" onClick={exportCurrentLiveData}>Експорт живих даних CRM</button>
          <button type="button" className="submit-button" onClick={createManualBackup}>Створити backup</button>
          <button type="button" onClick={() => backupInputRef.current?.click()}>Імпортувати backup</button>
        </div>
        <input
          ref={backupInputRef}
          type="file"
          accept=".json,application/json"
          className="backup-hidden-input"
          onChange={async (event) => {
            const file = event.target.files?.[0];
            if (file) await importBackupFile(file);
            event.currentTarget.value = '';
          }}
        />
        <div className="task-list">
          <Task icon={<Archive />} title="Експорт живих даних" text="Вивантажує поточні замовлення, клієнтів, оплати, документи, склад і аудит прямо з браузерного сховища у JSON." />
          <Task icon={<ShieldCheck />} title="Автобекап" text="CRM створює копію кожні 12 годин і тримає історію до 30 днів." />
          <Task icon={<Archive />} title="Що зберігається" text="Замовлення, клієнти, платежі, склад, закупки, аудит і договори." />
          <Task icon={<Truck />} title="Зовнішня копія" text="JSON-файл можна експортувати і зберегти у Google Drive, Dropbox або інше хмарне сховище." />
        </div>
        <div className="table stock-table">
          <div className="table-head">
            <span>Дата</span>
            <span>Тип</span>
            <span>Хто</span>
            <span>Дія</span>
          </div>
          {backups.map((backup) => (
            <div key={backup.id} className="table-row">
              <span>{backup.createdAt}</span>
              <span>{backup.source === 'auto' ? 'Авто' : backup.source === 'manual' ? 'Ручна' : 'Імпорт'}</span>
              <span>{backup.createdBy}</span>
              <span className="backup-action-cell">
                <button type="button" onClick={() => restoreBackup(backup.id)}>Відновити</button>
                <button type="button" onClick={() => exportBackup(backup.id)}>Експорт</button>
              </span>
            </div>
          ))}
          {backups.length === 0 && <div className="empty-state">Резервних копій поки немає.</div>}
        </div>
      </section>
      <section className="content-split">
        <div className="panel">
          <div className="panel-heading"><h2>Ролі та права</h2><span>{Object.keys(rolePermissions).length} ролей</span></div>
          <div className="task-list">
            {baseEmployeeRoles.map((role) => <Task key={role} icon={<ShieldCheck />} title={role} text={rolePermissions[role].join(', ')} />)}
          </div>
        </div>
        <div className="panel">
          <div className="panel-heading"><h2>Авторизація і журнал</h2><span>{actionLogs.length} записів</span></div>
          <div className="task-list">
            <Task icon={<ShieldCheck />} title="Логін і пароль" text="У backend-версії буде перевірка логіна, пароля, сесій і відновлення доступу." />
            <Task icon={<CheckCircle2 />} title="2FA" text="Увімкнено для керівника й адміністратора в моделі ролей." />
            <Task icon={<History />} title="Журнал дій" text="Критичні дії фіксують користувача, роль, документ, дату й коментар." />
          </div>
        </div>
      </section>
      <section className="content-split">
        <div className="panel">
          <div className="panel-heading"><h2>Доступи по модулях</h2><span>мінімально необхідний доступ</span></div>
          <div className="task-list">
            {baseEmployeeRoles.map((role) => <Task key={role} icon={<LayoutDashboard />} title={role} text={rolePageAccess[role].map((pageId) => navItems.find((item) => item.id === pageId)?.label ?? pageId).join(', ')} />)}
          </div>
        </div>
        <div className="panel">
          <div className="panel-heading"><h2>Точкові права</h2><span>{activeUser.role}</span></div>
          <div className="task-list">
            <Task icon={<ShieldCheck />} title="Поточна роль" text={`${activeUser.name}: ${activeUser.role}. 2FA: ${activeUser.twoFactor ? 'увімкнено' : 'не увімкнено'}.`} />
            <Task icon={<CheckCircle2 />} title="Дозволено" text={roleFinePermissions[activeUser.role].join(', ')} />
            <Task icon={<X />} title="Заборонено без окремого права" text="Видалення проведених документів, зміна старих залишків, зміна ціни після продажу, редагування закритого ремонту, зміна оплати заднім числом, списання без причини, повернення без посилання." />
          </div>
        </div>
      </section>
    </div>
  );
}

function Metric({ icon, label, value, hint }: { icon: React.ReactNode; label: string; value: string; hint: string }) {
  return <article className="metric-card"><div className="metric-icon">{icon}</div><span>{label}</span><strong>{value}</strong><p>{hint}</p></article>;
}

function MiniBars({ rows, formatValue = String }: { rows: Array<{ label: string; value: number }>; formatValue?: (value: number) => string }) {
  const max = Math.max(...rows.map((row) => row.value), 1);
  return (
    <div className="mini-bars">
      {rows.map((row) => (
        <div className="mini-bar" key={row.label}>
          <div className="mini-bar-top"><span>{row.label}</span><strong>{formatValue(row.value)}</strong></div>
          <div className="mini-bar-track"><span style={{ width: `${Math.max(8, (row.value / max) * 100)}%` }} /></div>
        </div>
      ))}
    </div>
  );
}

function Task({ icon, title, text }: { icon: React.ReactNode; title: string; text: string }) {
  return <article className="task"><div className="task-icon">{icon}</div><div><strong>{title}</strong><p>{text}</p></div></article>;
}

function Info({ label, value }: { label: string; value: string }) {
  return <div className="info-box"><span>{label}</span><strong>{value}</strong></div>;
}

function orderPriority(order: ServiceOrder) {
  if (daysSince(order.promisedDate) > 0) return 'Високий';
  if (order.status === 'Очікує запчастину' || daysSince(order.statusChangedAt) > 2) return 'Середній';
  return 'Звичайний';
}

function OrderSummaryList({ orders, selectedId, onSelect, showPriority = false, users = [] }: { orders: ServiceOrder[]; selectedId?: string; onSelect?: (id: string) => void; showPriority?: boolean; users?: User[] }) {
  return (
    <div className="order-list">
      {orders.map((order) => (
        <button type="button" className={selectedId === order.id ? 'order-card order-card-active' : 'order-card'} key={order.id} onClick={() => onSelect?.(order.id)}>
          <div className="order-main">
            <div>
              <div className="order-title"><strong>{order.id}</strong><span className={statusClass[order.status]}>{statusDisplay(order.status)}</span></div>
              <small>Етап: {statusStage(order.status)}</small>
              <h3>{order.device}</h3>
              <p>{order.client} · {order.issue}</p>
              <p>Менеджер: {getOrderCreatorName(order, users)}</p>
            </div>
            <div className="order-price"><strong>{showPriority ? orderPriority(order) : order.parts.length}</strong><span>{showPriority ? 'пріоритет' : 'запчастин'}</span></div>
          </div>
          <div className="order-meta"><span>{order.engineer}</span><span>{order.locationCode ?? 'без ячейки'}</span><span>{order.serial}</span>{showPriority && <span>Строк: {order.promisedDate}</span>}</div>
        </button>
      ))}
    </div>
  );
}



