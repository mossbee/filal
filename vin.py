def countVerticalPaths(cost, edge_nodes, edge_from, edge_to, k):
    # cost: the array representing the value of each node.
    # edge_nodes: number of nodes in the tree.
    # edge_from: integer array where the ith integer denotes one endpoint of the ith edge.
    # edge_to: integer array where the ith integer denotes the other endpoint of the ith edge
    # k an integer
    total_valid_paths = 0
    vertices_num = len(edge_from)
    for node in range(1, edge_nodes + 1):
        if cost[node] % k == 0:
            total_valid_paths += 1
        else:
            for 