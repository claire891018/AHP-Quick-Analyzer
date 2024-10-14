import numpy as np

def calculate_geometric_mean(matrix):
    geometric_means = np.prod(matrix, axis=1) ** (1 / matrix.shape[0])
    return geometric_means

def normalize_weights(geometric_means):
    weights = geometric_means / np.sum(geometric_means)
    return weights

def calculate_weights(matrix):
    geometric_means = calculate_geometric_mean(matrix)
    weights = normalize_weights(geometric_means)
    return weights
