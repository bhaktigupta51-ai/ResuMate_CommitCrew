import os
import math
import random
import csv
import zipfile
import xml.etree.ElementTree as ET

os.makedirs('results', exist_ok=True)
os.makedirs('charts', exist_ok=True)

# ----------------------------------------------------
# 1. XLSX Parser & Writer in Pure Python
# ----------------------------------------------------
def parse_xlsx(filename):
    with zipfile.ZipFile(filename, 'r') as z:
        shared_strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            tree = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for elem in tree.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t'):
                shared_strings.append(elem.text if elem.text else '')
        
        sheet_xml = z.read('xl/worksheets/sheet1.xml')
        tree = ET.fromstring(sheet_xml)
        rows_dict = {}
        for r_elem in tree.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row'):
            r_idx = int(r_elem.attrib['r'])
            row_vals = {}
            for cell in r_elem.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
                r_ref = cell.attrib.get('r')
                col_letter = ''.join([c for c in r_ref if c.isalpha()])
                t = cell.attrib.get('t')
                v = cell.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                val = v.text if v is not None else ''
                if t == 's' and val != '' and val.isdigit():
                    val = shared_strings[int(val)]
                row_vals[col_letter] = val
            rows_dict[r_idx] = row_vals
            
        sorted_r_keys = sorted(rows_dict.keys())
        cols = set()
        for k in sorted_r_keys:
            for col in rows_dict[k].keys():
                cols.add(col)
                
        def col2num(col):
            num = 0
            for c in col:
                num = num * 26 + (ord(c.upper()) - ord('A')) + 1
            return num - 1
            
        cols_sorted = sorted(list(cols), key=col2num)
        
        table = []
        for k in sorted_r_keys:
            r = rows_dict[k]
            row_lst = [r.get(c, '') for c in cols_sorted]
            table.append(row_lst)
        return table

def write_xlsx(filename, headers, rows):
    # Create simple openxml excel zip file
    shared_strings = []
    string_map = {}
    
    def get_ss_idx(s):
        s = str(s)
        if s not in string_map:
            string_map[s] = len(shared_strings)
            shared_strings.append(s)
        return string_map[s]

    def num2col(n):
        col = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            col = chr(65 + remainder) + col
        return col

    # build worksheet xml
    sheet_lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<sheetData>'
    ]
    
    # header row
    r_xml = '<row r="1">'
    for c_idx, h in enumerate(headers, start=1):
        s_idx = get_ss_idx(h)
        col_ref = num2col(c_idx) + "1"
        r_xml += f'<c r="{col_ref}" t="s"><v>{s_idx}</v></c>'
    r_xml += '</row>'
    sheet_lines.append(r_xml)
    
    for r_idx, row in enumerate(rows, start=2):
        r_xml = f'<row r="{r_idx}">'
        for c_idx, val in enumerate(row, start=1):
            col_ref = num2col(c_idx) + str(r_idx)
            try:
                fval = float(val)
                # check if integer
                if fval == int(fval):
                    r_xml += f'<c r="{col_ref}"><v>{int(fval)}</v></c>'
                else:
                    r_xml += f'<c r="{col_ref}"><v>{fval}</v></c>'
            except (ValueError, TypeError):
                s_idx = get_ss_idx(str(val))
                r_xml += f'<c r="{col_ref}" t="s"><v>{s_idx}</v></c>'
        r_xml += '</row>'
        sheet_lines.append(r_xml)
        
    sheet_lines.append('</sheetData></worksheet>')
    sheet_xml = '\n'.join(sheet_lines)

    # ss xml
    ss_lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{len(shared_strings)}" uniqueCount="{len(shared_strings)}">'
    ]
    for s in shared_strings:
        # escape xml
        s_esc = s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        ss_lines.append(f'<si><t>{s_esc}</t></si>')
    ss_lines.append('</sst>')
    ss_xml = '\n'.join(ss_lines)

    content_types = '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>'
    rels = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    wb_rels = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>'
    workbook = '<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></sheets></workbook>'

    with zipfile.ZipFile(filename, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', content_types)
        z.writestr('_rels/.rels', rels)
        z.writestr('xl/_rels/workbook.xml.rels', wb_rels)
        z.writestr('xl/workbook.xml', workbook)
        z.writestr('xl/sharedStrings.xml', ss_xml)
        z.writestr('xl/worksheets/sheet1.xml', sheet_xml)

# ----------------------------------------------------
# 2. Math & Machine Learning Utilities
# ----------------------------------------------------
def sigmoid(z):
    if z < -40: return 1e-15
    if z > 40: return 1.0 - 1e-15
    return 1.0 / (1.0 + math.exp(-z))

def mean(vals):
    return sum(vals) / len(vals) if vals else 0.0

def std_dev(vals):
    m = mean(vals)
    var = sum((x - m)**2 for x in vals) / (len(vals) - 1) if len(vals) > 1 else 0.0
    return math.sqrt(var)

def compute_auc(y_true, y_scores):
    pos = [s for t, s in zip(y_true, y_scores) if t == 1]
    neg = [s for t, s in zip(y_true, y_scores) if t == 0]
    if not pos or not neg: return 0.5
    wins = 0.0
    for p in pos:
        for n in neg:
            if p > n: wins += 1.0
            elif p == n: wins += 0.5
    return wins / (len(pos) * len(neg))

def compute_accuracy(y_true, y_preds):
    correct = sum(1 for t, p in zip(y_true, y_preds) if t == p)
    return correct / len(y_true) if y_true else 0.0

def mat_inv(A):
    n = len(A)
    AI = [A[i][:] + [1.0 if i == j else 0.0 for j in range(n)] for i in range(n)]
    for i in range(n):
        pivot = AI[i][i]
        if abs(pivot) < 1e-12:
            for k in range(i+1, n):
                if abs(AI[k][i]) > abs(pivot):
                    AI[i], AI[k] = AI[k], AI[i]
                    pivot = AI[i][i]
                    break
        for j in range(2*n):
            AI[i][j] /= pivot
        for k in range(n):
            if k != i:
                factor = AI[k][i]
                for j in range(2*n):
                    AI[k][j] -= factor * AI[i][j]
    return [row[n:] for row in AI]

class LogisticRegressionPure:
    def __init__(self, max_iter=10, tol=1e-6):
        self.max_iter = max_iter
        self.tol = tol
        self.weights = []
        self.bias = 0.0
        
    def fit(self, X, y):
        n_samples = len(X)
        n_features = len(X[0])
        # Design matrix with intercept column
        X_design = [[1.0] + row for row in X]
        p_dim = n_features + 1
        beta = [0.0] * p_dim
        
        l2 = 1e-4
        for _ in range(self.max_iter):
            probs = [sigmoid(sum(X_design[i][j] * beta[j] for j in range(p_dim))) for i in range(n_samples)]
            # Gradient: X^T * (y - p) - l2 * beta
            grad = [0.0] * p_dim
            for i in range(n_samples):
                err = y[i] - probs[i]
                for j in range(p_dim):
                    grad[j] += X_design[i][j] * err
            for j in range(1, p_dim):
                grad[j] -= l2 * beta[j]
                
            # Hessian: X^T * W * X + l2 * I
            H = [[0.0]*p_dim for _ in range(p_dim)]
            for i in range(n_samples):
                w = probs[i] * (1.0 - probs[i])
                w = max(w, 1e-6)
                row = X_design[i]
                for j in range(p_dim):
                    for k in range(p_dim):
                        H[j][k] += w * row[j] * row[k]
            for j in range(1, p_dim):
                H[j][j] += l2
                
            try:
                H_inv = mat_inv(H)
            except Exception:
                break
                
            delta = [sum(H_inv[j][k] * grad[k] for k in range(p_dim)) for j in range(p_dim)]
            beta = [beta[j] + delta[j] for j in range(p_dim)]
            if math.sqrt(sum(d**2 for d in delta)) < self.tol:
                break
                
        self.bias = beta[0]
        self.weights = beta[1:]
            
    def predict_proba(self, X):
        probs = []
        for i in range(len(X)):
            z = sum(X[i][j] * self.weights[j] for j in range(len(self.weights))) + self.bias
            probs.append(sigmoid(z))
        return probs

    def get_se(self, X, y):
        n_samples = len(X)
        n_features = len(X[0])
        X_design = [[1.0] + row for row in X]
        p_dim = n_features + 1
        H = [[0.0]*p_dim for _ in range(p_dim)]
        probs = self.predict_proba(X)
        for i in range(n_samples):
            w = probs[i] * (1.0 - probs[i])
            w = max(w, 1e-5)
            row = X_design[i]
            for j in range(p_dim):
                for k in range(p_dim):
                    H[j][k] += w * row[j] * row[k]
        try:
            cov = mat_inv(H)
            se = [math.sqrt(max(cov[j][j], 1e-8)) for j in range(1, p_dim)]
        except Exception:
            se = [0.2] * n_features
        return se

class TreeNode:
    def __init__(self, feature=None, threshold=None, left=None, right=None, value=None, n_samples=0):
        self.feature = feature
        self.threshold = threshold
        self.left = left
        self.right = right
        self.value = value
        self.n_samples = n_samples

class DecisionTreePure:
    def __init__(self, max_depth=3):
        self.max_depth = max_depth
        self.root = None
        
    def _gini(self, y):
        if not y: return 0.0
        p = sum(y) / len(y)
        return 1.0 - (p**2 + (1.0 - p)**2)
        
    def _build_tree(self, X, y, depth=0):
        n_samples = len(X)
        if n_samples == 0:
            return TreeNode(value=0.5, n_samples=0)
        p = sum(y) / n_samples
        if depth >= self.max_depth or len(set(y)) == 1 or n_samples < 4:
            return TreeNode(value=p, n_samples=n_samples)
            
        n_features = len(X[0])
        best_gini = 1.0
        best_feat, best_thresh = None, None
        best_left_X, best_left_y, best_right_X, best_right_y = [], [], [], []
        
        for f in range(n_features):
            vals = sorted(list(set(X[i][f] for i in range(n_samples))))
            thresholds = [(vals[i] + vals[i+1])/2.0 for i in range(len(vals)-1)]
            for t in thresholds:
                left_X, left_y, right_X, right_y = [], [], [], []
                for i in range(n_samples):
                    if X[i][f] <= t:
                        left_X.append(X[i]); left_y.append(y[i])
                    else:
                        right_X.append(X[i]); right_y.append(y[i])
                if not left_y or not right_y: continue
                g_left = self._gini(left_y)
                g_right = self._gini(right_y)
                w_gini = (len(left_y)/n_samples)*g_left + (len(right_y)/n_samples)*g_right
                if w_gini < best_gini:
                    best_gini = w_gini
                    best_feat = f
                    best_thresh = t
                    best_left_X, best_left_y = left_X, left_y
                    best_right_X, best_right_y = right_X, right_y
                    
        if best_feat is None:
            return TreeNode(value=p, n_samples=n_samples)
            
        left_child = self._build_tree(best_left_X, best_left_y, depth+1)
        right_child = self._build_tree(best_right_X, best_right_y, depth+1)
        return TreeNode(feature=best_feat, threshold=best_thresh, left=left_child, right=right_child, value=p, n_samples=n_samples)

    def fit(self, X, y):
        self.root = self._build_tree(X, y, 0)
        
    def _predict_row(self, node, row):
        if node.feature is None:
            return node.value
        if row[node.feature] <= node.threshold:
            return self._predict_row(node.left, row)
        else:
            return self._predict_row(node.right, row)
            
    def predict_proba(self, X):
        return [self._predict_row(self.root, row) for row in X]

    def export_text_rules(self, node=None, depth=0, feature_names=None):
        if node is None: node = self.root
        indent = "|   " * depth
        if node.feature is None:
            return f"{indent}|--- class/probability: {node.value:.3f} (samples: {node.n_samples})\n"
        feat_name = feature_names[node.feature] if feature_names else f"feature_{node.feature}"
        res = f"{indent}|--- {feat_name} <= {node.threshold:.3f}\n"
        res += self.export_text_rules(node.left, depth+1, feature_names)
        res += f"{indent}|--- {feat_name} >  {node.threshold:.3f}\n"
        res += self.export_text_rules(node.right, depth+1, feature_names)
        return res

class RandomForestPure:
    def __init__(self, n_estimators=60, max_depth=3, seed=42):
        self.n_estimators = n_estimators
        self.max_depth = max_depth
        self.trees = []
        self.seed = seed
        
    def fit(self, X, y):
        rng = random.Random(self.seed)
        n_samples = len(X)
        self.trees = []
        for _ in range(self.n_estimators):
            boot_idx = [rng.randint(0, n_samples-1) for _ in range(n_samples)]
            boot_X = [X[i] for i in boot_idx]
            boot_y = [y[i] for i in boot_idx]
            tree = DecisionTreePure(max_depth=self.max_depth)
            tree.fit(boot_X, boot_y)
            self.trees.append(tree)
            
    def predict_proba(self, X):
        all_probs = [tree.predict_proba(X) for tree in self.trees]
        n_samples = len(X)
        final_probs = []
        for i in range(n_samples):
            p_mean = sum(all_probs[t][i] for t in range(len(self.trees))) / len(self.trees)
            final_probs.append(p_mean)
        return final_probs

# ----------------------------------------------------
# 3. Repeated Stratified 5-Fold Generator
# ----------------------------------------------------
def get_repeated_stratified_kfold_indices(y, n_splits=5, n_repeats=20, seed=42):
    rng = random.Random(seed)
    pos_idx = [i for i, val in enumerate(y) if val == 1]
    neg_idx = [i for i, val in enumerate(y) if val == 0]
    
    splits = []
    for r in range(n_repeats):
        rng.shuffle(pos_idx)
        rng.shuffle(neg_idx)
        
        pos_folds = [pos_idx[i::n_splits] for i in range(n_splits)]
        neg_folds = [neg_idx[i::n_splits] for i in range(n_splits)]
        
        for f in range(n_splits):
            test_idx = pos_folds[f] + neg_folds[f]
            train_idx = [i for i in range(len(y)) if i not in test_idx]
            splits.append((train_idx, test_idx))
    return splits

def assign_band(prob):
    if prob >= 0.7:
        return 'likely high'
    elif prob >= 0.4:
        return 'borderline'
    else:
        return 'needs development'

# SVG Generator for Crisp Visualizations
def save_svg_roc(filename, title, models_roc):
    # models_roc: dict of name -> (y_true, probs, auc_score)
    svg = ['<svg xmlns="http://www.w3.org/2000/svg" width="700" height="500" viewBox="0 0 700 500" style="background-color:#ffffff; font-family:Arial, sans-serif;">']
    svg.append(f'<text x="350" y="35" text-anchor="middle" font-size="18" font-weight="bold" fill="#222222">{title}</text>')
    
    # Axes box
    x0, y0, w, h = 90, 70, 520, 360
    svg.append(f'<rect x="{x0}" y="{y0}" width="{w}" height="{h}" fill="#fcfcfc" stroke="#cccccc" stroke-width="1"/>')
    
    # Grid lines
    for i in range(1, 5):
        gx = x0 + i * (w / 5.0)
        gy = y0 + i * (h / 5.0)
        svg.append(f'<line x1="{gx}" y1="{y0}" x2="{gx}" y2="{y0+h}" stroke="#e0e0e0" stroke-dasharray="4,4"/>')
        svg.append(f'<line x1="{x0}" y1="{gy}" x2="{x0+w}" y2="{gy}" stroke="#e0e0e0" stroke-dasharray="4,4"/>')
        
        # Axis labels
        svg.append(f'<text x="{gx}" y="{y0+h+20}" text-anchor="middle" font-size="11" fill="#555">{i*0.2:.1f}</text>')
        svg.append(f'<text x="{x0-12}" y="{y0+h - i*(h/5.0) + 4}" text-anchor="end" font-size="11" fill="#555">{i*0.2:.1f}</text>')
        
    svg.append(f'<text x="{x0}" y="{y0+h+20}" text-anchor="middle" font-size="11" fill="#555">0.0</text>')
    svg.append(f'<text x="{x0+w}" y="{y0+h+20}" text-anchor="middle" font-size="11" fill="#555">1.0</text>')
    svg.append(f'<text x="{x0-12}" y="{y0+h+4}" text-anchor="end" font-size="11" fill="#555">0.0</text>')
    svg.append(f'<text x="{x0-12}" y="{y0+4}" text-anchor="end" font-size="11" fill="#555">1.0</text>')
    
    svg.append(f'<text x="{x0 + w/2}" y="{y0+h+45}" text-anchor="middle" font-size="13" font-weight="bold" fill="#333">False Positive Rate</text>')
    svg.append(f'<text x="30" y="{y0 + h/2}" text-anchor="middle" font-size="13" font-weight="bold" fill="#333" transform="rotate(-90, 30, {y0+h/2})">True Positive Rate</text>')

    # Random chance line
    svg.append(f'<line x1="{x0}" y1="{y0+h}" x2="{x0+w}" y2="{y0}" stroke="#888888" stroke-dasharray="6,6" stroke-width="1.5"/>')
    
    colors = {'Logistic Regression': '#1f77b4', 'Decision Tree (depth 3)': '#ff7f0e', 'Random Forest': '#2ca02c'}
    leg_y = y0 + h - 120
    
    for name, (y_true, probs, auc_val) in models_roc.items():
        color = colors.get(name, '#333333')
        # Compute ROC points
        thresholds = sorted(list(set(probs)), reverse=True)
        pts = []
        n_pos = sum(y_true)
        n_neg = len(y_true) - n_pos
        
        for t in [1.1] + thresholds + [-0.1]:
            tp = sum(1 for y_i, p_i in zip(y_true, probs) if y_i == 1 and p_i >= t)
            fp = sum(1 for y_i, p_i in zip(y_true, probs) if y_i == 0 and p_i >= t)
            tpr = tp / n_pos if n_pos > 0 else 0
            fpr = fp / n_neg if n_neg > 0 else 0
            
            px = x0 + fpr * w
            py = y0 + h - tpr * h
            pts.append(f'{px:.1f},{py:.1f}')
            
        svg.append(f'<polyline points="{" ".join(pts)}" fill="none" stroke="{color}" stroke-width="3"/>')
        
        # Legend
        svg.append(f'<line x1="{x0+w-210}" y1="{leg_y}" x2="{x0+w-180}" y2="{leg_y}" stroke="{color}" stroke-width="3"/>')
        svg.append(f'<text x="{x0+w-170}" y="{leg_y+4}" font-size="12" font-weight="bold" fill="#333">{name} (AUC = {auc_val:.3f})</text>')
        leg_y += 25
        
    svg.append('</svg>')
    with open(filename, 'w', encoding='utf-8') as f:
        f.write('\n'.join(svg))

# ----------------------------------------------------
# 4. Main Dataset Processing Function
# ----------------------------------------------------
def run_dataset_analysis(excel_path, dataset_name, target_keyword):
    print(f"\n=======================================================")
    print(f" ANALYZING: {dataset_name} ({excel_path})")
    print(f"=======================================================")
    
    raw_data = parse_xlsx(excel_path)
    raw_headers = raw_data[0]
    data_rows = raw_data[1:]
    
    # Strip spaces from column names
    headers = [str(h).strip() for h in raw_headers]
    print(f"Cleaned Headers: {headers}")
    
    # Identify id column and target column
    id_col_idx = 0
    target_col_idx = len(headers) - 1
    
    for idx, h in enumerate(headers):
        if h.lower() == 'id':
            id_col_idx = idx
        elif target_keyword.lower() in h.lower():
            target_col_idx = idx
            
    id_col_name = headers[id_col_idx]
    target_col_name = headers[target_col_idx]
    
    predictor_indices = [i for i in range(len(headers)) if i not in (id_col_idx, target_col_idx)]
    predictor_names = [headers[i] for i in predictor_indices]
    
    print(f"ID Column: '{id_col_name}' (idx {id_col_idx})")
    print(f"Target Column: '{target_col_name}' (idx {target_col_idx})")
    print(f"Predictor Columns ({len(predictor_names)}): {predictor_names}")
    
    # Filter data rows (remove empty trailing rows where id or target is blank)
    valid_rows = []
    for r in data_rows:
        if len(r) > max(id_col_idx, target_col_idx) and str(r[id_col_idx]).strip() != '' and str(r[target_col_idx]).strip() != '':
            valid_rows.append(r)
    data_rows = valid_rows

    # Extract arrays
    ids = [r[id_col_idx] for r in data_rows]
    y = [int(float(r[target_col_idx])) for r in data_rows]
    X_raw = [[float(r[i]) for i in predictor_indices] for r in data_rows]
    n_samples = len(X_raw)
    n_features = len(predictor_names)
    
    print(f"Total Rows Kept: {n_samples}")
    
    # Standardize features (Z-score)
    means = [mean([X_raw[i][j] for i in range(n_samples)]) for j in range(n_features)]
    sds = [std_dev([X_raw[i][j] for i in range(n_samples)]) for j in range(n_features)]
    
    X_std = [[(X_raw[i][j] - means[j]) / sds[j] for j in range(n_features)] for i in range(n_samples)]
    
    # Repeated Stratified 5-Fold Cross-Validation (20 repeats = 100 splits)
    splits = get_repeated_stratified_kfold_indices(y, n_splits=5, n_repeats=20, seed=42)
    
    models = {
        'Logistic Regression': 'lr',
        'Decision Tree (depth 3)': 'dt',
        'Random Forest': 'rf'
    }
    
    cv_metrics = {m: {'auc': [], 'acc': []} for m in models}
    oof_accum = {m: [0.0]*n_samples for m in models}
    
    print("\nRunning Repeated Stratified 5-Fold CV (20 repeats = 100 iterations)...")
    for train_idx, test_idx in splits:
        X_train_raw = [X_raw[i] for i in train_idx]
        X_train_std = [X_std[i] for i in train_idx]
        y_train = [y[i] for i in train_idx]
        
        X_test_raw = [X_raw[i] for i in test_idx]
        X_test_std = [X_std[i] for i in test_idx]
        y_test = [y[i] for i in test_idx]
        
        # 1. Logistic Regression
        lr = LogisticRegressionPure(max_iter=10)
        lr.fit(X_train_std, y_train)
        probs_lr = lr.predict_proba(X_test_std)
        preds_lr = [1 if p >= 0.5 else 0 for p in probs_lr]
        cv_metrics['Logistic Regression']['auc'].append(compute_auc(y_test, probs_lr))
        cv_metrics['Logistic Regression']['acc'].append(compute_accuracy(y_test, preds_lr))
        for idx_pos, i_orig in enumerate(test_idx):
            oof_accum['Logistic Regression'][i_orig] += probs_lr[idx_pos] / 20.0
            
        # 2. Decision Tree depth 3
        dt = DecisionTreePure(max_depth=3)
        dt.fit(X_train_raw, y_train)
        probs_dt = dt.predict_proba(X_test_raw)
        preds_dt = [1 if p >= 0.5 else 0 for p in probs_dt]
        cv_metrics['Decision Tree (depth 3)']['auc'].append(compute_auc(y_test, probs_dt))
        cv_metrics['Decision Tree (depth 3)']['acc'].append(compute_accuracy(y_test, preds_dt))
        for idx_pos, i_orig in enumerate(test_idx):
            oof_accum['Decision Tree (depth 3)'][i_orig] += probs_dt[idx_pos] / 20.0
            
        # 3. Random Forest
        rf = RandomForestPure(n_estimators=50, max_depth=3, seed=42)
        rf.fit(X_train_raw, y_train)
        probs_rf = rf.predict_proba(X_test_raw)
        preds_rf = [1 if p >= 0.5 else 0 for p in probs_rf]
        cv_metrics['Random Forest']['auc'].append(compute_auc(y_test, probs_rf))
        cv_metrics['Random Forest']['acc'].append(compute_accuracy(y_test, preds_rf))
        for idx_pos, i_orig in enumerate(test_idx):
            oof_accum['Random Forest'][i_orig] += probs_rf[idx_pos] / 20.0

    print("\n-------------------------------------------------------")
    print(f" PERFORMANCE METRICS ({dataset_name}):")
    print("-------------------------------------------------------")
    for name in models:
        mean_auc = mean(cv_metrics[name]['auc'])
        std_auc = std_dev(cv_metrics[name]['auc'])
        mean_acc = mean(cv_metrics[name]['acc'])
        std_acc = std_dev(cv_metrics[name]['acc'])
        print(f"{name:25s} | Mean AUC: {mean_auc:.4f} ± {std_auc:.4f} | Mean Accuracy: {mean_acc:.4f} ± {std_acc:.4f}")

    # Full Logistic Regression on Standardized Features for Odds Ratios per +1 SD
    lr_full = LogisticRegressionPure(max_iter=10)
    lr_full.fit(X_std, y)
    se_std = lr_full.get_se(X_std, y)
    
    print("\n-------------------------------------------------------")
    print(f" LOGISTIC REGRESSION ODDS RATIOS PER +1 SD:")
    print("-------------------------------------------------------")
    print(f"{'Predictor':35s} | {'Mean':>7s} | {'SD':>7s} | {'Coeff (+1SD)':>12s} | {'Odds Ratio':>10s} | {'95% CI':>20s}")
    print("-" * 105)
    
    or_records = []
    for j in range(n_features):
        b = lr_full.weights[j]
        se = se_std[j]
        or_val = math.exp(b)
        ci_low = math.exp(b - 1.96 * se)
        ci_high = math.exp(b + 1.96 * se)
        p_name = predictor_names[j]
        m_val = means[j]
        sd_val = sds[j]
        print(f"{p_name:35s} | {m_val:7.2f} | {sd_val:7.2f} | {b:12.4f} | {or_val:10.4f} | [{ci_low:.4f}, {ci_high:.4f}]")
        or_records.append({'Predictor': p_name, 'Mean': m_val, 'SD': sd_val, 'Coeff': b, 'OR': or_val, 'CI_low': ci_low, 'CI_high': ci_high})

    # Full Decision Tree (depth 3) for Rules
    dt_full = DecisionTreePure(max_depth=3)
    dt_full.fit(X_raw, y)
    tree_rules = dt_full.export_text_rules(feature_names=predictor_names)
    
    print("\n-------------------------------------------------------")
    print(f" DEPTH-3 DECISION TREE RULES ({dataset_name}):")
    print("-------------------------------------------------------")
    print(tree_rules)
    
    # Save tree rules to text file
    rules_file = f"results/{dataset_name.lower().replace(' ', '_')}_tree_rules.txt"
    with open(rules_file, 'w', encoding='utf-8') as f:
        f.write(f"DEPTH-3 DECISION TREE RULES - {dataset_name}\n\n")
        f.write(tree_rules)

    # Person-level Results File
    lr_probs = oof_accum['Logistic Regression']
    results_rows = []
    
    band_counts = {'likely high': 0, 'borderline': 0, 'needs development': 0}
    
    for i in range(n_samples):
        person_id = ids[i]
        prob = round(lr_probs[i], 4)
        band = assign_band(prob)
        band_counts[band] += 1
        results_rows.append([person_id, prob, band])
        
    res_prefix = dataset_name.lower().replace(' ', '_')
    csv_file = f"results/{res_prefix}_results.csv"
    xlsx_file = f"results/{res_prefix}_results.xlsx"
    
    # Write CSV
    with open(csv_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['id', 'predicted_probability', 'band'])
        writer.writerows(results_rows)
        
    # Write XLSX
    write_xlsx(xlsx_file, ['id', 'predicted_probability', 'band'], results_rows)
    
    print("\n-------------------------------------------------------")
    print(f" PERSON-LEVEL BAND DISTRIBUTION ({dataset_name}):")
    print("-------------------------------------------------------")
    for band, count in band_counts.items():
        pct = (count / n_samples) * 100
        print(f"  {band:20s}: {count:4d} persons ({pct:.1f}%)")
    print(f"Saved individual results to '{csv_file}' and '{xlsx_file}'")

    # Generate Crisp SVG ROC Curve Chart
    roc_chart_file = f"charts/{res_prefix}_roc_curves.svg"
    models_roc_data = {
        name: (y, oof_accum[name], mean(cv_metrics[name]['auc']))
        for name in models
    }
    save_svg_roc(roc_chart_file, f"ROC Curves (5-Fold CV, 20 Repeats) - {dataset_name}", models_roc_data)
    print(f"Saved ROC curve chart to '{roc_chart_file}'")
    
    return {
        'dataset_name': dataset_name,
        'cv_metrics': cv_metrics,
        'or_records': or_records,
        'tree_rules': tree_rules,
        'band_counts': band_counts,
        'results_rows': results_rows
    }

if __name__ == '__main__':
    jds = run_dataset_analysis('JDS Skill Traits.xlsx', 'JDS Skill Traits', 'salary_hike')
    sds = run_dataset_analysis('SDS Personality Traits.xlsx', 'SDS Personality Traits', 'success')
    print("\n=======================================================")
    print(" ALL ANALYSES AND RESULTS COMPLETED SUCCESSFULLY!")
    print("=======================================================")
