import os
import sys
import math
import random
import csv
import zipfile
import xml.etree.ElementTree as ET

os.makedirs('charts', exist_ok=True)
os.makedirs('results', exist_ok=True)

# Attempt importing standard ML libraries (pandas, sklearn, statsmodels, matplotlib)
try:
    import numpy as np
    import pandas as pd
    import matplotlib.pyplot as plt
    import seaborn as sns
    from sklearn.model_selection import RepeatedStratifiedKFold
    from sklearn.linear_model import LogisticRegression
    from sklearn.tree import DecisionTreeClassifier, export_text, plot_tree
    from sklearn.ensemble import RandomForestClassifier
    from sklearn.metrics import roc_auc_score, accuracy_score, roc_curve
    from sklearn.preprocessing import StandardScaler
    import statsmodels.api as sm
    USE_SKLEARN = True
except ImportError:
    USE_SKLEARN = False

def assign_band(prob):
    if prob >= 0.7:
        return 'likely high'
    elif prob >= 0.4:
        return 'borderline'
    else:
        return 'needs development'

def process_dataset_sklearn(filepath, dataset_name, target_hint):
    print(f"\n==================================================")
    print(f" PROCESSING DATASET: {dataset_name} ({filepath}) [Scikit-Learn Pipeline]")
    print(f"==================================================")
    
    df = pd.read_excel(filepath)
    df.columns = df.columns.astype(str).str.strip()
    
    id_col = 'id' if 'id' in df.columns else df.columns[0]
    target_col = None
    for col in df.columns:
        if target_hint.lower() in col.lower():
            target_col = col
            break
    if target_col is None:
        target_col = df.columns[-1]
        
    df = df.dropna(subset=[id_col, target_col]).copy()
    df[target_col] = df[target_col].astype(int)
    
    predictor_cols = [col for col in df.columns if col not in [id_col, target_col]]
    
    print(f"Total Rows Kept: {len(df)}")
    print(f"ID Column: '{id_col}'")
    print(f"Target Column: '{target_col}'")
    print(f"Predictors ({len(predictor_cols)}): {predictor_cols}")
    
    X = df[predictor_cols].copy()
    y = df[target_col].copy()
    ids = df[id_col].copy()
    
    # Repeated Stratified 5-Fold CV (20 repeats = 100 splits)
    rskf = RepeatedStratifiedKFold(n_splits=5, n_repeats=20, random_state=42)
    models = {
        'Logistic Regression': LogisticRegression(max_iter=1000, random_state=42),
        'Decision Tree (depth 3)': DecisionTreeClassifier(max_depth=3, random_state=42),
        'Random Forest': RandomForestClassifier(n_estimators=100, max_depth=3, random_state=42)
    }
    
    cv_results = {}
    oof_predictions = {name: np.zeros(len(df)) for name in models}
    
    for name, model in models.items():
        auc_scores, acc_scores = [], []
        prob_accum = np.zeros(len(df))
        
        for train_idx, test_idx in rskf.split(X, y):
            X_tr, X_te = X.iloc[train_idx], X.iloc[test_idx]
            y_tr, y_te = y.iloc[train_idx], y.iloc[test_idx]
            
            model.fit(X_tr, y_tr)
            probs = model.predict_proba(X_te)[:, 1]
            preds = model.predict(X_te)
            
            auc_scores.append(roc_auc_score(y_te, probs))
            acc_scores.append(accuracy_score(y_te, preds))
            prob_accum[test_idx] += probs / 20.0
            
        cv_results[name] = {
            'mean_auc': np.mean(auc_scores), 'std_auc': np.std(auc_scores),
            'mean_acc': np.mean(acc_scores), 'std_acc': np.std(acc_scores)
        }
        oof_predictions[name] = prob_accum
        
        print(f"\n--- {name} CV Metrics (20 repeats x 5 folds = 100 runs) ---")
        print(f"Mean AUC:      {cv_results[name]['mean_auc']:.4f} ± {cv_results[name]['std_auc']:.4f}")
        print(f"Mean Accuracy: {cv_results[name]['mean_acc']:.4f} ± {cv_results[name]['std_acc']:.4f}")

    # Logistic Regression Odds Ratios per +1 SD
    scaler = StandardScaler()
    X_scaled = pd.DataFrame(scaler.fit_transform(X), columns=predictor_cols)
    X_scaled_const = sm.add_constant(X_scaled)
    logit_mod = sm.Logit(y, X_scaled_const)
    logit_res = logit_mod.fit(disp=False)
    
    params = logit_res.params[1:]
    conf = logit_res.conf_int().iloc[1:]
    se = logit_res.bse[1:]
    
    or_df = pd.DataFrame({
        'Predictor': predictor_cols,
        'Mean': X.mean().values,
        'SD': X.std().values,
        'Coeff (+1 SD)': params.values,
        'Odds Ratio': np.exp(params.values),
        '95% CI Lower': np.exp(conf[0].values),
        '95% CI Upper': np.exp(conf[1].values)
    })
    print(f"\n--- Logistic Regression Odds Ratios per +1 SD ---")
    print(or_df.to_string(index=False))

    # Decision Tree Rules
    dt_model = DecisionTreeClassifier(max_depth=3, random_state=42)
    dt_model.fit(X, y)
    tree_rules = export_text(dt_model, feature_names=predictor_cols)
    print(f"\n--- Depth-3 Decision Tree Rules ---")
    print(tree_rules)
    
    rules_file = f"results/{dataset_name.lower().replace(' ', '_')}_tree_rules.txt"
    with open(rules_file, 'w', encoding='utf-8') as f:
        f.write(f"DEPTH-3 DECISION TREE RULES - {dataset_name}\n\n" + tree_rules)

    # Results CSV and Excel
    lr_probs = oof_predictions['Logistic Regression']
    results_df = pd.DataFrame({
        'id': ids,
        'predicted_probability': np.round(lr_probs, 4),
        'band': [assign_band(p) for p in lr_probs]
    })
    
    res_prefix = dataset_name.lower().replace(' ', '_')
    csv_file = f"results/{res_prefix}_results.csv"
    xlsx_file = f"results/{res_prefix}_results.xlsx"
    results_df.to_csv(csv_file, index=False)
    results_df.to_excel(xlsx_file, index=False)
    print(f"\nSaved results to '{csv_file}' and '{xlsx_file}'")

    # ROC Curves
    plt.figure(figsize=(8, 6))
    for name in models:
        fpr, tpr, _ = roc_curve(y, oof_predictions[name])
        plt.plot(fpr, tpr, lw=2, label=f"{name} (AUC = {cv_results[name]['mean_auc']:.3f})")
    plt.plot([0, 1], [0, 1], 'k--', lw=1.5, label='Random Chance')
    plt.xlabel('False Positive Rate'); plt.ylabel('True Positive Rate')
    plt.title(f'ROC Curves - {dataset_name}')
    plt.legend(loc='lower right')
    plt.tight_layout()
    plt.savefig(f"charts/{res_prefix}_roc_curves.png", dpi=300)
    plt.close()

# --------------------------------------------------------------------
# 2. PURE PYTHON FALLBACK PIPELINE (ZERO EXTERNAL DEPENDENCIES)
# --------------------------------------------------------------------
def parse_xlsx_pure(filename):
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
            
        sorted_keys = sorted(rows_dict.keys())
        cols = set(col for k in sorted_keys for col in rows_dict[k].keys())
        cols_sorted = sorted(list(cols), key=lambda c: sum((ord(char)-64)*(26**i) for i, char in enumerate(reversed(c))))
        
        table = []
        for k in sorted_keys:
            r = rows_dict[k]
            row_lst = [r.get(c, '') for c in cols_sorted]
            table.append(row_lst)
        return table

def sigmoid(z):
    if z < -40: return 1e-15
    if z > 40: return 1.0 - 1e-15
    return 1.0 / (1.0 + math.exp(-z))

def mean_val(vals):
    return sum(vals) / len(vals) if vals else 0.0

def std_val(vals):
    m = mean_val(vals)
    var = sum((x - m)**2 for x in vals) / (len(vals) - 1) if len(vals) > 1 else 0.0
    return math.sqrt(var)

def compute_auc_pure(y_true, y_scores):
    pos = [s for t, s in zip(y_true, y_scores) if t == 1]
    neg = [s for t, s in zip(y_true, y_scores) if t == 0]
    if not pos or not neg: return 0.5
    wins = sum(1.0 if p > n else (0.5 if p == n else 0.0) for p in pos for n in neg)
    return wins / (len(pos) * len(neg))

def mat_inv_pure(A):
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
        X_design = [[1.0] + row for row in X]
        p_dim = n_features + 1
        beta = [0.0] * p_dim
        l2 = 1e-4
        
        for _ in range(self.max_iter):
            probs = [sigmoid(sum(X_design[i][j] * beta[j] for j in range(p_dim))) for i in range(n_samples)]
            grad = [0.0] * p_dim
            for i in range(n_samples):
                err = y[i] - probs[i]
                for j in range(p_dim):
                    grad[j] += X_design[i][j] * err
            for j in range(1, p_dim):
                grad[j] -= l2 * beta[j]
                
            H = [[0.0]*p_dim for _ in range(p_dim)]
            for i in range(n_samples):
                w = max(probs[i] * (1.0 - probs[i]), 1e-6)
                row = X_design[i]
                for j in range(p_dim):
                    for k in range(p_dim):
                        H[j][k] += w * row[j] * row[k]
            for j in range(1, p_dim):
                H[j][j] += l2
                
            try:
                H_inv = mat_inv_pure(H)
            except Exception:
                break
                
            delta = [sum(H_inv[j][k] * grad[k] for k in range(p_dim)) for j in range(p_dim)]
            beta = [beta[j] + delta[j] for j in range(p_dim)]
            if math.sqrt(sum(d**2 for d in delta)) < self.tol:
                break
                
        self.bias = beta[0]
        self.weights = beta[1:]
        
    def predict_proba(self, X):
        return [sigmoid(sum(X[i][j] * self.weights[j] for j in range(len(self.weights))) + self.bias) for i in range(len(X))]

    def get_se(self, X, y):
        n_samples = len(X)
        n_features = len(X[0])
        X_design = [[1.0] + row for row in X]
        p_dim = n_features + 1
        H = [[0.0]*p_dim for _ in range(p_dim)]
        probs = self.predict_proba(X)
        for i in range(n_samples):
            w = max(probs[i] * (1.0 - probs[i]), 1e-5)
            row = X_design[i]
            for j in range(p_dim):
                for k in range(p_dim):
                    H[j][k] += w * row[j] * row[k]
        try:
            cov = mat_inv_pure(H)
            return [math.sqrt(max(cov[j][j], 1e-8)) for j in range(1, p_dim)]
        except Exception:
            return [0.2] * n_features

class TreeNodePure:
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
        if n_samples == 0: return TreeNodePure(value=0.5, n_samples=0)
        p = sum(y) / n_samples
        if depth >= self.max_depth or len(set(y)) == 1 or n_samples < 4:
            return TreeNodePure(value=p, n_samples=n_samples)
            
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
                w_gini = (len(left_y)/n_samples)*self._gini(left_y) + (len(right_y)/n_samples)*self._gini(right_y)
                if w_gini < best_gini:
                    best_gini = w_gini
                    best_feat, best_thresh = f, t
                    best_left_X, best_left_y = left_X, left_y
                    best_right_X, best_right_y = right_X, right_y
                    
        if best_feat is None:
            return TreeNodePure(value=p, n_samples=n_samples)
            
        left_child = self._build_tree(best_left_X, best_left_y, depth+1)
        right_child = self._build_tree(best_right_X, best_right_y, depth+1)
        return TreeNodePure(feature=best_feat, threshold=best_thresh, left=left_child, right=right_child, value=p, n_samples=n_samples)

    def fit(self, X, y):
        self.root = self._build_tree(X, y, 0)
        
    def _predict_row(self, node, row):
        if node.feature is None: return node.value
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
    def __init__(self, n_estimators=50, max_depth=3, seed=42):
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
        return [sum(all_probs[t][i] for t in range(len(self.trees))) / len(self.trees) for i in range(len(X))]

def get_repeated_stratified_kfold_indices_pure(y, n_splits=5, n_repeats=20, seed=42):
    rng = random.Random(seed)
    pos_idx = [i for i, val in enumerate(y) if val == 1]
    neg_idx = [i for i, val in enumerate(y) if val == 0]
    splits = []
    for _ in range(n_repeats):
        rng.shuffle(pos_idx); rng.shuffle(neg_idx)
        pos_folds = [pos_idx[i::n_splits] for i in range(n_splits)]
        neg_folds = [neg_idx[i::n_splits] for i in range(n_splits)]
        for f in range(n_splits):
            test_idx = pos_folds[f] + neg_folds[f]
            train_idx = [i for i in range(len(y)) if i not in test_idx]
            splits.append((train_idx, test_idx))
    return splits

def process_dataset_pure(filepath, dataset_name, target_hint):
    print(f"\n==================================================")
    print(f" PROCESSING DATASET: {dataset_name} ({filepath}) [Pure Python Pipeline]")
    print(f"==================================================")

    raw_data = parse_xlsx_pure(filepath)
    raw_headers = raw_data[0]
    raw_rows = raw_data[1:]
    headers = [str(h).strip() for h in raw_headers]
    
    id_col_idx = 0
    target_col_idx = len(headers) - 1
    for idx, h in enumerate(headers):
        if h.lower() == 'id': id_col_idx = idx
        elif target_hint.lower() in h.lower(): target_col_idx = idx
            
    id_col_name = headers[id_col_idx]
    target_col_name = headers[target_col_idx]
    predictor_indices = [i for i in range(len(headers)) if i not in (id_col_idx, target_col_idx)]
    predictor_names = [headers[i] for i in predictor_indices]

    valid_rows = []
    for r in raw_rows:
        if len(r) > max(id_col_idx, target_col_idx) and str(r[id_col_idx]).strip() != '' and str(r[target_col_idx]).strip() != '':
            valid_rows.append(r)

    ids = [r[id_col_idx] for r in valid_rows]
    y = [int(float(r[target_col_idx])) for r in valid_rows]
    X_raw = [[float(r[i]) for i in predictor_indices] for r in valid_rows]
    n_samples = len(valid_rows)

    print(f"Total Rows Kept: {n_samples}")
    print(f"ID Column: '{id_col_name}'")
    print(f"Target Column: '{target_col_name}'")
    print(f"Predictors ({len(predictor_names)}): {predictor_names}")

    means = [mean_val([X_raw[i][j] for i in range(n_samples)]) for j in range(len(predictor_names))]
    sds = [std_val([X_raw[i][j] for i in range(n_samples)]) for j in range(len(predictor_names))]
    X_std = [[(X_raw[i][j] - means[j]) / sds[j] for j in range(len(predictor_names))] for i in range(n_samples)]

    splits = get_repeated_stratified_kfold_indices_pure(y, n_splits=5, n_repeats=20, seed=42)
    models = {'Logistic Regression': 'lr', 'Decision Tree (depth 3)': 'dt', 'Random Forest': 'rf'}
    
    cv_metrics = {m: {'auc': [], 'acc': []} for m in models}
    oof_accum = {m: [0.0]*n_samples for m in models}

    for train_idx, test_idx in splits:
        X_tr_std = [X_std[i] for i in train_idx]; X_te_std = [X_std[i] for i in test_idx]
        X_tr_raw = [X_raw[i] for i in train_idx]; X_te_raw = [X_raw[i] for i in test_idx]
        y_tr = [y[i] for i in train_idx]; y_te = [y[i] for i in test_idx]

        lr = LogisticRegressionPure(max_iter=10)
        lr.fit(X_tr_std, y_tr)
        p_lr = lr.predict_proba(X_te_std)
        cv_metrics['Logistic Regression']['auc'].append(compute_auc_pure(y_te, p_lr))
        cv_metrics['Logistic Regression']['acc'].append(mean_val([1 if (p>=0.5)==yt else 0 for p, yt in zip(p_lr, y_te)]))
        for idx_pos, i_orig in enumerate(test_idx): oof_accum['Logistic Regression'][i_orig] += p_lr[idx_pos] / 20.0

        dt = DecisionTreePure(max_depth=3)
        dt.fit(X_tr_raw, y_tr)
        p_dt = dt.predict_proba(X_te_raw)
        cv_metrics['Decision Tree (depth 3)']['auc'].append(compute_auc_pure(y_te, p_dt))
        cv_metrics['Decision Tree (depth 3)']['acc'].append(mean_val([1 if (p>=0.5)==yt else 0 for p, yt in zip(p_dt, y_te)]))
        for idx_pos, i_orig in enumerate(test_idx): oof_accum['Decision Tree (depth 3)'][i_orig] += p_dt[idx_pos] / 20.0

        rf = RandomForestPure(n_estimators=40, max_depth=3, seed=42)
        rf.fit(X_tr_raw, y_tr)
        p_rf = rf.predict_proba(X_te_raw)
        cv_metrics['Random Forest']['auc'].append(compute_auc_pure(y_te, p_rf))
        cv_metrics['Random Forest']['acc'].append(mean_val([1 if (p>=0.5)==yt else 0 for p, yt in zip(p_rf, y_te)]))
        for idx_pos, i_orig in enumerate(test_idx): oof_accum['Random Forest'][i_orig] += p_rf[idx_pos] / 20.0

    print("\n--- Model Cross-Validation Metrics (20 repeats x 5 folds = 100 runs) ---")
    for name in models:
        m_auc = mean_val(cv_metrics[name]['auc']); s_auc = std_val(cv_metrics[name]['auc'])
        m_acc = mean_val(cv_metrics[name]['acc']); s_acc = std_val(cv_metrics[name]['acc'])
        print(f"{name:25s} | Mean AUC: {m_auc:.4f} ± {s_auc:.4f} | Mean Accuracy: {m_acc:.4f} ± {s_acc:.4f}")

    lr_full = LogisticRegressionPure(max_iter=10)
    lr_full.fit(X_std, y)
    se_std = lr_full.get_se(X_std, y)

    print(f"\n--- Logistic Regression Odds Ratios per +1 SD ---")
    print(f"{'Predictor':35s} | {'Mean':>7s} | {'SD':>7s} | {'Coeff (+1SD)':>12s} | {'Odds Ratio':>10s} | {'95% CI':>20s}")
    print("-" * 105)
    for j, p_name in enumerate(predictor_names):
        b = lr_full.weights[j]
        se_j = se_std[j]
        or_val = math.exp(b)
        ci_low = math.exp(b - 1.96 * se_j)
        ci_high = math.exp(b + 1.96 * se_j)
        print(f"{p_name:35s} | {means[j]:7.2f} | {sds[j]:7.2f} | {b:12.4f} | {or_val:10.4f} | [{ci_low:.4f}, {ci_high:.4f}]")

    dt_full = DecisionTreePure(max_depth=3)
    dt_full.fit(X_raw, y)
    tree_rules = dt_full.export_text_rules(feature_names=predictor_names)
    print(f"\n--- Depth-3 Decision Tree Rules ---")
    print(tree_rules)

    rules_file = f"results/{dataset_name.lower().replace(' ', '_')}_tree_rules.txt"
    with open(rules_file, 'w', encoding='utf-8') as f:
        f.write(f"DEPTH-3 DECISION TREE RULES - {dataset_name}\n\n" + tree_rules)

    lr_probs = oof_accum['Logistic Regression']
    results_rows = [[ids[i], round(lr_probs[i], 4), assign_band(lr_probs[i])] for i in range(n_samples)]
    res_prefix = dataset_name.lower().replace(' ', '_')
    csv_file = f"results/{res_prefix}_results.csv"
    with open(csv_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['id', 'predicted_probability', 'band'])
        writer.writerows(results_rows)

    print(f"\nSaved results to '{csv_file}'")

# Dispatch function
def process_dataset(filepath, dataset_name, target_hint):
    if USE_SKLEARN:
        return process_dataset_sklearn(filepath, dataset_name, target_hint)
    else:
        return process_dataset_pure(filepath, dataset_name, target_hint)

if __name__ == '__main__':
    process_dataset('JDS Skill Traits.xlsx', 'JDS Skill Traits', 'salary_hike')
    process_dataset('SDS Personality Traits.xlsx', 'SDS Personality Traits', 'success')
    print("\nProcessing complete! All analyses, models, tree rules, and results saved successfully.")
