def values(fn):
    def _style(ct, ws, idx):
        r = idx['r']
        c = idx['c']
        label = idx['label']
        ws.write(r, c, label)
    return _style
