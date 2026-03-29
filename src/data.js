// src/data.js
export const initialNodes = [
  {
    id: '1',
    type: 'productionNode', // Our custom component
    data: { 
      label: 'M2 Gimbal Assembly', 
      metrics: {
        cycleTime: 4.5,    // Days
        training: 0.85,   // 85% coverage
        ecns: 0,          // Open Engineering Changes
        docs: 'Up-to-date'
      },
      documentation: {
        wiName: 'Gimbal Assembly',
        wiNumber: 'WI-5000',
        hasChecklist: true
      },
    },
    position: { x: 250, y: 0 },
  },
  {
    id: '2',
    type: 'productionNode',
    data: { 
      label: 'M2 AZ Head', 
      metrics: {
        cycleTime: 6.2, 
        training: 0.4, 
        ecns: 3, 
        docs: 'Draft'
      },
      documentation: {
        wiName: 'AZ Head Assembly',
        wiNumber: 'WI-4000',
        hasChecklist: null
      },
    },
    position: { x: 0, y: 150 },
  },
  {
    id: '3',
    type: 'productionNode',
    data: { 
      label: 'M2 RL Arm', 
      metrics: {
        cycleTime: 6.2, 
        training: 0.4, 
        ecns: 3, 
        docs: 'Draft'
      },
      documentation: {
        wiName: 'RL Arm Assembly',
        wiNumber: 'WI-3000',
        hasChecklist: true
      },
    },
    position: { x: 0, y: 150 },
  },
  {
    id: '4',
    type: 'productionNode',
    data: { 
      label: 'Payload', 
      metrics: {
        cycleTime: 6.2, 
        training: 0.4, 
        ecns: 3, 
        docs: 'Draft'
      },
      documentation: {
        wiName: 'Payload Assembly',
        wiNumber: 'WI-2000',
        hasChecklist: true
      },
    },
    position: { x: 0, y: 150 },
  },
  {
    id: '5',
    type: 'productionNode',
    data: { 
      label: 'CJ18', 
      metrics: {
        cycleTime: 6.2, 
        training: 0.4, 
        ecns: 3, 
        docs: 'Draft'
      },
      documentation: {
        wiName: 'CJ18 assembly',
        wiNumber: 'WI-1000',
        hasChecklist: true
      },
    },
    position: { x: 0, y: 150 },
  },
];

export const initialEdges = [
  { id: 'e1-2', source: '1', target: '2', animated: false },
  { id: 'e1-3', source: '1', target: '3', animated: false },
  { id: 'e1-4', source: '1', target: '4', animated: false },
  { id: 'e2-3', source: '4', target: '5', animated: false },
];