import numpy as np

def calculate_consistency_ratio(matrix, weights):
    """計算一致性比率 (C.R.) 和一致性指數 (C.I.)"""
    n = matrix.shape[0]
    # λ_max
    lambda_max = np.sum(np.sum(matrix * weights, axis=0) / weights) / n
    # C.I.
    consistency_index = (lambda_max - n) / (n - 1)
    # R.I.
    random_index = 0.9
    # C.R. = C.I. / R.I.
    consistency_ratio = consistency_index / random_index

    print(f"λ_max: {lambda_max}")
    print(f"C.I.: {consistency_index}")
    print(f"C.R.: {consistency_ratio}")

    return consistency_index, consistency_ratio

