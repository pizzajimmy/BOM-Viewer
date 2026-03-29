import React, { useState, useEffect } from 'react';
import { ReactFlow, useNodesState, useEdgesState, Background, Controls } from '@xyflow/react';
import '@xyflow/react/dist/style.css';
import ProductionNode from './ProductionNode';
import { initialNodes as nodesDef, initialEdges as edgesDef } from './data'; // Renamed consistently
import { getLayoutedElements } from './layout';

const nodeTypes = { productionNode: ProductionNode };

export default function App() {
  // 1. Run the layout engine ONCE before setting initial state
  const { nodes: layoutedNodes, edges: layoutedEdges } = getLayoutedElements(
    nodesDef,
    edgesDef
  );
  
  const [nodes, setNodes, onNodesChange] = useNodesState(layoutedNodes);
  const [edges, setEdges, onEdgesChange] = useEdgesState(layoutedEdges);
  const [view, setView] = useState('CYCLE_TIME');
  
  // 2. Update nodes whenever the view changes
  useEffect(() => {
    setNodes((nds) =>
      nds.map((node) => ({
        ...node,
        data: { ...node.data, view },
      }))
    );
  }, [view, setNodes]);

  return (
    <div style={{ width: '100vw', height: '100vh', background: '#1a1a1a' }}>
      {/* Control Panel */}
      <div style={{ 
        position: 'absolute', 
        zIndex: 10, 
        top: 20, 
        left: 20, 
        display: 'flex', 
        gap: '10px',
        background: 'rgba(255,255,255,0.1)',
        padding: '10px',
        borderRadius: '8px',
        backdropFilter: 'blur(4px)'
      }}>
        <button onClick={() => setView('CYCLE_TIME')}>Cycle Time View</button>
        <button onClick={() => setView('TRAINING')}>Training Coverage</button>
        <button onClick={() => setView('ENGINEERING')}>Engineering Changes</button>
        <button onClick={() => setView('DOCS')}>Documentation View</button>
      </div>
      
      <ReactFlow
        nodes={nodes}
        edges={edges}
        onNodesChange={onNodesChange}
        nodeTypes={nodeTypes}
        fitView
      >
        <Background color="#333" gap={20} />
        <Controls />
      </ReactFlow>
    </div>
  );
}