import itertools as it


def styles(ct):
    columns = ct.levels.columns
    index = ct.levels.index
    values = ct.levels.values
    x = []
    y = []

    matrix = lambda idx, i: idx[:i + 1] + [idx[i]] * \
        (len(idx) - len(idx[:i + 1]))

    if columns:
        for i in range(len(columns)):
            for val in values:
                col = matrix(columns, i)
                x.append(tuple(col + [val]))
    else:
        x = [(x,) for x in values]

    for i in range(len(index)):
        idx = matrix(index, i)
        y.append(tuple(idx))

    styles_matrix = []
    for z in zip(*x):
          styles_matrix.append(tuple([''] * len(ct.levels.index)) + \
          tuple([''] * len(ct.levels.values)) + z)
    
    styles_matrix.append(tuple([''] * len(ct.levels.index)) + \
          tuple([''] * len(ct.levels.values)) + \
          tuple([''] * len(x)))
    
    for i in y:
        val = []
        for c in x:
            val.append(tuple([i, c]))
        styles_matrix.append(tuple(i) + tuple([''] * len(values)) + tuple(val))
    
    return styles_matrix


def values(fn):
    def _style(ct, ws, idx):
        r = idx['r']
        c = idx['c']
        label = idx['label']
        ws.write(r, c, label)
    return _style
