import { Handle, Position } from '@xyflow/react';

export default function ProductionNode({ data }) {
  // Logic to determine color based on the 'view' passed in data
  const { metrics, view, documentation } = data;

  const renderDocView = () => {
    if (!documentation) {
      return (
        <div style={{ color: '#fff', background: '#cf1322', padding: '4px', borderRadius: '4px', textAlign: 'center', fontWeight: 'bold' }}>
          ⚠️ NO WI ASSIGNED
        </div>
      );
    }

    return (
      <div style={{ fontSize: '10px', color: '#333' }}>
        <div style={{ fontWeight: 'bold' }}>{documentation.wiNumber}</div>
        <div>{documentation.wiName}</div>
        {documentation.hasChecklist ? (
          <div style={{ marginTop: '5px', color: '#0958d9', display: 'flex', alignItems: 'center', gap: '4px' }}>
            📋 Checklist Integrated
          </div>
        ) : (
          <div style={{ marginTop: '5px', color: '#d48806' }}>
            ⚠️ No Checklist
          </div>
        )}
      </div>
    );
  };
  
  const getStatusColor = () => {
    if (view === 'DOCS') {
      if (!documentation) return '#ffa39e'; // Soft red
      return documentation.hasChecklist ? '#b7eb8f' : '#ffe58f'; // Green vs Yellow
    }

    switch (view) {
      case 'CYCLE_TIME':
        return metrics.cycleTime > 5 ? '#ff4d4f' : '#73d13d'; // Red if > Takt time (5 days)
      case 'TRAINING':
        return metrics.training < 0.5 ? '#ff4d4f' : metrics.training < 0.8 ? '#ffec3d' : '#73d13d';
      case 'ENGINEERING':
        return metrics.ecns > 0 ? '#ffa940' : '#73d13d';
      default:
        return '#fff';
    }
  };

  return (
    <div style={{ 
      padding: '10px', 
      borderRadius: '5px', 
      background: getStatusColor(),
      border: '1px solid #555',
      minWidth: '150px'
    }}>
      <Handle type="target" position={Position.Top} />
      <div style={{ fontWeight: 'bold', fontSize: '12px' }}>{data.label}</div>
      <div style={{ fontSize: '10px', marginTop: '5px' }}>
        {view === 'CYCLE_TIME' && `Time: ${metrics.cycleTime}d`}
        {view === 'TRAINING' && `Skills: ${Math.round(metrics.training * 100)}%`}
        {view === 'ENGINEERING' && `ECNs: ${metrics.ecns}`}
        {view === 'DOCS' && renderDocView()}
      </div>
      <Handle type="source" position={Position.Bottom} />
    </div>
  );
}