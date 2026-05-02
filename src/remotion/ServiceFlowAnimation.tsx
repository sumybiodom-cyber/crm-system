import { AbsoluteFill, interpolate, spring, useCurrentFrame, useVideoConfig } from 'remotion';

const steps = ['Прийом', 'Діагностика', 'Ремонт', 'Оплата', 'Видача'];

export function ServiceFlowAnimation() {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  const orbit = interpolate(frame % 180, [0, 180], [0, 360]);
  const pulse = spring({ frame: frame % 60, fps, config: { damping: 12, stiffness: 90 } });

  return (
    <AbsoluteFill style={{ background: '#f8fbfa', fontFamily: 'Inter, Arial, sans-serif', overflow: 'hidden' }}>
      <div
        style={{
          position: 'absolute',
          inset: 24,
          border: '1px solid #d8e4df',
          borderRadius: 8,
          background: 'linear-gradient(135deg, #ffffff 0%, #edf7f3 100%)',
        }}
      />
      <div
        style={{
          position: 'absolute',
          width: 190,
          height: 190,
          right: 56,
          top: 42,
          borderRadius: '50%',
          border: '16px solid #d8ece5',
          transform: `rotate(${orbit}deg)`,
        }}
      >
        <div style={{ width: 22, height: 22, borderRadius: '50%', background: '#00a878', marginLeft: 78, marginTop: -18 }} />
      </div>
      <div style={{ position: 'absolute', left: 54, top: 54, color: '#18342c' }}>
        <div style={{ fontSize: 18, color: '#417267', fontWeight: 700 }}>Маршрут замовлення</div>
        <div style={{ marginTop: 10, fontSize: 42, lineHeight: 1.05, fontWeight: 800, maxWidth: 360 }}>
          Все видно з першого погляду
        </div>
      </div>
      <div style={{ position: 'absolute', left: 54, right: 54, bottom: 58, display: 'flex', gap: 12 }}>
        {steps.map((step, index) => {
          const start = index * 24;
          const progress = interpolate((frame + 20) % 160, [start, start + 18], [0, 1], {
            extrapolateLeft: 'clamp',
            extrapolateRight: 'clamp',
          });
          return (
            <div
              key={step}
              style={{
                flex: 1,
                minHeight: 82,
                borderRadius: 8,
                border: '1px solid #cfe0da',
                background: progress > 0.65 ? '#18342c' : '#ffffff',
                color: progress > 0.65 ? '#ffffff' : '#18342c',
                padding: 16,
                transform: `translateY(${(1 - progress) * 12}px) scale(${0.96 + progress * 0.04})`,
                opacity: 0.55 + progress * 0.45,
              }}
            >
              <div style={{ fontSize: 22, fontWeight: 800 }}>{index + 1}</div>
              <div style={{ marginTop: 8, fontSize: 17, fontWeight: 700 }}>{step}</div>
            </div>
          );
        })}
      </div>
      <div
        style={{
          position: 'absolute',
          right: 96,
          top: 128,
          width: 92 + pulse * 12,
          height: 92 + pulse * 12,
          borderRadius: 8,
          background: '#00a878',
          color: '#ffffff',
          display: 'grid',
          placeItems: 'center',
          fontSize: 30,
          fontWeight: 900,
          boxShadow: '0 20px 45px rgba(0, 120, 86, 0.22)',
        }}
      >
        CRM
      </div>
    </AbsoluteFill>
  );
}
