import dagre from 'dagre';

const dagreGraph = new dagre.graphlib.Graph();
dagreGraph.setDefaultEdgeLabel(() => ({}));

const nodeWidth = 200;
const nodeHeight = 80;

export const getLayoutedElements = (nodes, edges, direction = 'LR') => {
  dagreGraph.setGraph({ rankdir: direction, ranksep: 200, nodesep: 100 });

  nodes.forEach((node) => {
    dagreGraph.setNode(node.id, { width: nodeWidth, height: nodeHeight });
  });

  edges.forEach((edge) => {
    dagreGraph.setEdge(edge.source, edge.target);
  });

  dagre.layout(dagreGraph);

  const newNodes = nodes.map((node) => {
    const nodeWithPosition = dagreGraph.node(node.id);
    return {
      ...node,
      position: {
        x: nodeWithPosition.x + nodeWidth / 2,
        y: nodeWithPosition.y + nodeHeight / 2,
      },
    };
  });

  return { nodes: newNodes, edges };
};