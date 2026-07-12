"""Data-validation rules engine (Qt-free).

Rules live in each master table's config under ``validations`` (a list of rule
dicts, see ``normalize_rule``). The engine compiles rules into per-row
predicates, keeps a live violation map for cell colouring, and supports
incremental re-validation after single-cell edits.

Two rule modes:
  - builder: structured when/then condition lists (Excel-like)
  - expr:    a restricted Python expression evaluated per row (AST whitelist,
             never uses eval/exec)

Scopes:
  - ""          → the master table itself (rows of tables[master])
  - "SubName"   → a sub table (rows of sub_tables["Master.SubName"])

Cross-table access:
  - sub-scope rules may reference master fields as ``master.Field``
  - master-scope rules may use aggregate conditions over a sub table
    (builder: cond {"agg": {...}}; expr: any_sub()/count_sub())
"""

import ast
import re
import uuid


# ──────────────────────────────────────────────────────────────────────────────
# Rule model
# ──────────────────────────────────────────────────────────────────────────────

SEVERITIES = ("error", "warn")

# operators usable in builder conditions (value2 only for "between")
OPS = ("eq", "ne", "empty", "not_empty", "contains", "not_contains",
       "in_list", "not_in", "gt", "ge", "lt", "le", "between", "regex")

OP_LABELS = {
    "eq": "等於", "ne": "不等於", "empty": "為空", "not_empty": "非空",
    "contains": "包含", "not_contains": "不包含",
    "in_list": "在清單中", "not_in": "不在清單中",
    "gt": ">", "ge": "≥", "lt": "<", "le": "≤",
    "between": "介於", "regex": "符合Regex",
}

COUNT_OP_LABELS = {"eq": "＝", "ne": "≠", "ge": "≥", "le": "≤", "gt": "＞", "lt": "＜"}

DEFAULT_COLOR = "#E5484D"


def new_rule(scope=""):
    return {
        "id": uuid.uuid4().hex[:8],
        "name": "新規則",
        "enabled": True,
        "severity": "error",
        "color": DEFAULT_COLOR,
        "scope": scope,
        "mode": "builder",
        "when": {"logic": "and", "conds": []},
        "then": [],
        "expr": "",
        "mark": [],
    }


def normalize_rule(d):
    """Fill defaults on a rule dict loaded from config (in place, returns it)."""
    d.setdefault("id", uuid.uuid4().hex[:8])
    d.setdefault("name", "未命名規則")
    d.setdefault("enabled", True)
    d.setdefault("severity", "error")
    d.setdefault("color", DEFAULT_COLOR)
    d.setdefault("scope", "")
    d.setdefault("mode", "builder")
    w = d.setdefault("when", {"logic": "and", "conds": []})
    w.setdefault("logic", "and")
    w.setdefault("conds", [])
    d.setdefault("then", [])
    d.setdefault("expr", "")
    d.setdefault("mark", [])
    if d["severity"] not in SEVERITIES:
        d["severity"] = "error"
    return d


# ──────────────────────────────────────────────────────────────────────────────
# Value helpers
# ──────────────────────────────────────────────────────────────────────────────

def is_empty(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def _to_num(v):
    if v is None:
        return None
    if isinstance(v, bool):
        return float(v)
    try:
        s = str(v).strip()
        if s == "":
            return None
        return float(s)
    except (ValueError, TypeError):
        return None


def typed_value(raw, col_type):
    """Coerce a DataFrame cell (usually str) by its configured column type.
    Blank numerics become None so `empty()` and comparisons behave sanely."""
    if col_type in ("int", "float"):
        return _to_num(raw)
    if col_type == "bool":
        if isinstance(raw, bool):
            return raw
        return str(raw).strip().lower() in ("true", "1", "yes")
    return "" if raw is None else str(raw)


def _as_bool(x):
    if isinstance(x, bool):
        return x
    return str(x).strip().lower() in ("true", "1", "yes")


def _eq(a, b):
    """Equality that treats numeric-looking values numerically ("1.0" == "1")
    and bools leniently (True == "true" == "1")."""
    if isinstance(a, bool) or isinstance(b, bool):
        return _as_bool(a) == _as_bool(b)
    na, nb = _to_num(a), _to_num(b)
    if na is not None and nb is not None:
        return na == nb
    return ("" if a is None else str(a).strip()) == \
           ("" if b is None else str(b).strip())


def _cmp_num(a, b, op):
    na, nb = _to_num(a), _to_num(b)
    if na is None or nb is None:
        return False
    if op == "gt":
        return na > nb
    if op == "ge":
        return na >= nb
    if op == "lt":
        return na < nb
    if op == "le":
        return na <= nb
    return False


def eval_op(val, op, value="", value2=""):
    """Evaluate one builder operator against a cell value. Returns bool."""
    if op == "empty":
        return is_empty(val)
    if op == "not_empty":
        return not is_empty(val)
    if op == "eq":
        return _eq(val, value)
    if op == "ne":
        return not _eq(val, value)
    if op == "contains":
        return str(value) in ("" if val is None else str(val))
    if op == "not_contains":
        return str(value) not in ("" if val is None else str(val))
    if op in ("in_list", "not_in"):
        items = [t.strip() for t in str(value).split(",") if t.strip() != ""]
        hit = any(_eq(val, t) for t in items)
        return hit if op == "in_list" else not hit
    if op in ("gt", "ge", "lt", "le"):
        return _cmp_num(val, value, op)
    if op == "between":
        lo, hi = _to_num(value), _to_num(value2)
        n = _to_num(val)
        if n is None or lo is None or hi is None:
            return False
        return lo <= n <= hi
    if op == "regex":
        try:
            return re.search(str(value), "" if val is None else str(val)) is not None
        except re.error:
            return False
    return False


# ──────────────────────────────────────────────────────────────────────────────
# Safe expression evaluator (AST whitelist — never eval/exec)
# ──────────────────────────────────────────────────────────────────────────────

class ExprError(Exception):
    pass


_ALLOWED_CMP = (ast.Eq, ast.NotEq, ast.Lt, ast.LtE, ast.Gt, ast.GtE,
                ast.In, ast.NotIn)
_ALLOWED_BIN = (ast.Add, ast.Sub, ast.Mult, ast.Div, ast.FloorDiv,
                ast.Mod, ast.Pow)

# functions callable inside expressions; row-context funcs are injected by the
# resolver (empty / num / match / any_sub / count_sub / master_val)
_SAFE_BUILTINS = {
    "len": len, "str": str, "int": int, "float": float,
    "abs": abs, "min": min, "max": max, "round": round,
}
_CTX_FUNCS = ("empty", "num", "match", "any_sub", "count_sub")
_ALLOWED_FUNC_NAMES = set(_SAFE_BUILTINS) | set(_CTX_FUNCS)


class SafeExpr:
    """Parse once, validate the AST against a whitelist, then interpret per row.

    Names resolve to current-row fields (typed by column config);
    ``master.Field`` resolves through the resolver; constants True/False/None
    work as usual.
    """

    def __init__(self, expr: str):
        self.expr = expr
        try:
            tree = ast.parse(expr, mode="eval")
        except SyntaxError as e:
            raise ExprError(f"語法錯誤：{e.msg}（位置 {e.offset}）")
        self._validate(tree)
        self._tree = tree

    @classmethod
    def _validate(cls, node):
        for n in ast.walk(node):
            if isinstance(n, (ast.Expression, ast.Constant, ast.Name, ast.Load,
                              ast.List, ast.Tuple, ast.IfExp,
                              ast.BoolOp, ast.And, ast.Or,
                              ast.UnaryOp, ast.Not, ast.USub, ast.UAdd,
                              ast.BinOp, ast.Compare, ast.Call, ast.keyword)):
                if isinstance(n, ast.Call):
                    if not isinstance(n.func, ast.Name) \
                            or n.func.id not in _ALLOWED_FUNC_NAMES:
                        raise ExprError("只允許呼叫："
                                        + ", ".join(sorted(_ALLOWED_FUNC_NAMES)))
                    if n.keywords:
                        raise ExprError("不支援關鍵字參數")
                continue
            if isinstance(n, ast.Attribute):
                if not (isinstance(n.value, ast.Name) and n.value.id == "master"
                        and isinstance(n.ctx, ast.Load)) \
                        or n.attr.startswith("_"):
                    raise ExprError("屬性存取只允許 master.欄位")
                continue
            if isinstance(n, tuple(_ALLOWED_CMP) + tuple(_ALLOWED_BIN)):
                continue
            raise ExprError(f"不允許的語法：{type(n).__name__}")

    def eval(self, resolver):
        return self._eval(self._tree.body, resolver)

    def _eval(self, n, r):
        if isinstance(n, ast.Constant):
            return n.value
        if isinstance(n, ast.Name):
            return r.get_name(n.id)
        if isinstance(n, ast.Attribute):        # master.Field (validated above)
            return r.get_master(n.attr)
        if isinstance(n, (ast.List, ast.Tuple)):
            return [self._eval(e, r) for e in n.elts]
        if isinstance(n, ast.BoolOp):
            if isinstance(n.op, ast.And):
                v = True
                for e in n.values:
                    v = self._eval(e, r)
                    if not v:
                        return v
                return v
            v = False
            for e in n.values:
                v = self._eval(e, r)
                if v:
                    return v
            return v
        if isinstance(n, ast.UnaryOp):
            v = self._eval(n.operand, r)
            if isinstance(n.op, ast.Not):
                return not v
            if isinstance(n.op, ast.USub):
                return -(v or 0)
            return +(v or 0)
        if isinstance(n, ast.BinOp):
            a, b = self._eval(n.left, r), self._eval(n.right, r)
            try:
                if isinstance(n.op, ast.Add):
                    return a + b
                if isinstance(n.op, ast.Sub):
                    return a - b
                if isinstance(n.op, ast.Mult):
                    return a * b
                if isinstance(n.op, ast.Div):
                    return a / b
                if isinstance(n.op, ast.FloorDiv):
                    return a // b
                if isinstance(n.op, ast.Mod):
                    return a % b
                if isinstance(n.op, ast.Pow):
                    return a ** b
            except (TypeError, ZeroDivisionError):
                return None
        if isinstance(n, ast.Compare):
            left = self._eval(n.left, r)
            for op, comp in zip(n.ops, n.comparators):
                right = self._eval(comp, r)
                if not self._compare(left, op, right):
                    return False
                left = right
            return True
        if isinstance(n, ast.IfExp):
            return self._eval(n.body, r) if self._eval(n.test, r) \
                else self._eval(n.orelse, r)
        if isinstance(n, ast.Call):
            fname = n.func.id
            args = [self._eval(a, r) for a in n.args]
            if fname in _SAFE_BUILTINS:
                try:
                    return _SAFE_BUILTINS[fname](*args)
                except (TypeError, ValueError):
                    return None
            return r.call(fname, args, n)      # context funcs
        raise ExprError(f"無法求值：{type(n).__name__}")

    @staticmethod
    def _compare(a, op, b):
        if isinstance(op, ast.Eq):
            return _eq(a, b) if not (a is None or b is None) else a is b or a == b
        if isinstance(op, ast.NotEq):
            return not SafeExpr._compare(a, ast.Eq(), b)
        if isinstance(op, ast.In):
            if isinstance(b, str):
                return ("" if a is None else str(a)) in b
            try:
                return any(_eq(a, x) for x in b)
            except TypeError:
                return False
        if isinstance(op, ast.NotIn):
            return not SafeExpr._compare(a, ast.In(), b)
        # ordering comparisons: numeric only; None never satisfies
        na, nb = _to_num(a), _to_num(b)
        if na is None or nb is None:
            # allow lexicographic compare when both are non-numeric strings
            if isinstance(a, str) and isinstance(b, str):
                if isinstance(op, ast.Lt):
                    return a < b
                if isinstance(op, ast.LtE):
                    return a <= b
                if isinstance(op, ast.Gt):
                    return a > b
                if isinstance(op, ast.GtE):
                    return a >= b
            return False
        if isinstance(op, ast.Lt):
            return na < nb
        if isinstance(op, ast.LtE):
            return na <= nb
        if isinstance(op, ast.Gt):
            return na > nb
        if isinstance(op, ast.GtE):
            return na >= nb
        return False


# ──────────────────────────────────────────────────────────────────────────────
# Row context / resolver
# ──────────────────────────────────────────────────────────────────────────────

class _RowCtx:
    """Field access for one row being validated.

    master scope: sheet == master table; get("Field") reads the master row;
                  any_sub/count_sub aggregate over the row's sub-table rows.
    sub scope:    sheet == "Master.Sub"; get("Field") reads the sub row;
                  get("master.Field") / master.Field reads the parent row.
    """

    def __init__(self, engine, master, scope, row_idx):
        self.engine = engine
        self.master = master
        self.scope = scope            # "" or sub name
        self.row_idx = row_idx
        self._master_row_idx = None   # lazily resolved for sub scope

    # ── builder access (dotted master.X allowed as field name) ──
    def get(self, field):
        if field.startswith("master."):
            return self.get_master(field[7:])
        return self._own(field)

    def _own(self, field):
        e = self.engine
        if self.scope:
            sheet = f"{self.master}.{self.scope}"
            df = e.manager.sub_tables.get(sheet)
            ctype = e.sub_col_type(self.master, self.scope, field)
        else:
            df = e.manager.tables.get(self.master)
            ctype = e.master_col_type(self.master, field)
        if df is None or field not in df.columns or self.row_idx not in df.index:
            return None
        return typed_value(df.at[self.row_idx, field], ctype)

    def get_master(self, field):
        e = self.engine
        if not self.scope:            # master scope: master.X == X
            return self._own(field)
        midx = self._resolve_master_idx()
        if midx is None:
            return None
        df = e.manager.tables.get(self.master)
        if df is None or field not in df.columns:
            return None
        return typed_value(df.at[midx, field], e.master_col_type(self.master, field))

    def _resolve_master_idx(self):
        if self._master_row_idx is not None:
            return self._master_row_idx
        self._master_row_idx = self.engine.master_idx_for_sub(
            self.master, self.scope, self.row_idx)
        return self._master_row_idx

    # ── expr context functions ──
    def get_name(self, name):
        if name == "master":
            raise ExprError("master 要用 master.欄位 形式")
        if name in ("True", "False", "None"):   # py<3.8 safety, normally Constant
            return {"True": True, "False": False, "None": None}[name]
        return self.get(name)

    def call(self, fname, args, node=None):
        if fname == "empty":
            return is_empty(args[0]) if args else True
        if fname == "num":
            return _to_num(args[0]) if args else None
        if fname == "match":
            if len(args) < 2:
                return False
            try:
                return re.search(str(args[0]),
                                 "" if args[1] is None else str(args[1])) is not None
            except re.error:
                return False
        if fname in ("any_sub", "count_sub"):
            return self._agg_func(fname, args)
        raise ExprError(f"未知函式 {fname}")

    def _agg_func(self, fname, args):
        if self.scope:
            raise ExprError("any_sub/count_sub 只能用在母表規則")
        if not args:
            raise ExprError(f"{fname}(子表名, 條件式?) 至少要有子表名")
        sub_name = str(args[0])
        cond_expr = str(args[1]) if len(args) > 1 and args[1] is not None else ""
        cnt = self.engine.count_sub_rows(self.master, self.row_idx,
                                         sub_name, cond_expr)
        return (cnt > 0) if fname == "any_sub" else cnt


# ──────────────────────────────────────────────────────────────────────────────
# Engine
# ──────────────────────────────────────────────────────────────────────────────

class ValidationEngine:
    def __init__(self, manager):
        self.manager = manager
        self.rules = {}          # {master: [rule, ...]} (normalized, enabled+disabled)
        self._compiled = {}      # {rule_id: SafeExpr or None}
        self._sub_exprs = {}     # {(rule_id_or_adhoc, expr_str): SafeExpr} cache
        self.violations = {}     # {(sheet, row_idx, col): set(rule_id)}
        self._row_cells = {}     # {(sheet, row_idx): set((sheet,row,col))}
        self.rule_by_id = {}     # {rule_id: (master, rule)}
        self._pk_index = {}      # {master: {pk_str: row_idx}}
        self._fk_index = {}      # {(master, sub): {fk_str: [row_idx,...]}}
        self.last_errors = []    # [(rule_name, message)] compile problems

    # ── config access ─────────────────────────────────────────────────────────

    def table_rules(self, master):
        return self.rules.get(master, [])

    def master_col_type(self, master, col):
        return (self.manager.config.get(master, {})
                .get("columns", {}).get(col, {}).get("type", "string"))

    def sub_col_type(self, master, sub, col):
        return (self.manager.config.get(master, {})
                .get("sub_tables", {}).get(sub, {})
                .get("columns", {}).get(col, {}).get("type", "string"))

    def master_pk(self, master):
        df = self.manager.tables.get(master)
        cfg = self.manager.config.get(master, {})
        pk = cfg.get("primary_key", "")
        if pk and df is not None and pk in df.columns:
            return pk
        if df is not None and len(df.columns):
            return df.columns[0]
        return ""

    def sub_fk(self, master, sub):
        fk = (self.manager.config.get(master, {})
              .get("sub_tables", {}).get(sub, {}).get("foreign_key", ""))
        return fk or self.master_pk(master)

    # ── (re)load & compile ────────────────────────────────────────────────────

    def reload(self, validate=True):
        """Re-read rules from config, recompile expressions, revalidate all."""
        self.rules = {}
        self._compiled = {}
        self._sub_exprs = {}
        self.rule_by_id = {}
        self.last_errors = []
        for master in self.manager.tables:
            lst = self.manager.config.get(master, {}).get("validations", [])
            rules = [normalize_rule(r) for r in lst if isinstance(r, dict)]
            self.rules[master] = rules
            for r in rules:
                self.rule_by_id[r["id"]] = (master, r)
                if r.get("mode") == "expr" and r.get("expr", "").strip():
                    try:
                        self._compiled[r["id"]] = SafeExpr(r["expr"])
                    except ExprError as e:
                        self._compiled[r["id"]] = None
                        self.last_errors.append((r.get("name", "?"), str(e)))
        if validate:
            self.validate_all()

    # ── index building ────────────────────────────────────────────────────────

    def _build_indexes(self, master):
        df = self.manager.tables.get(master)
        pk = self.master_pk(master)
        idx = {}
        if df is not None and pk and pk in df.columns:
            for ri, v in df[pk].items():
                idx.setdefault(str(v).strip(), ri)
        self._pk_index[master] = idx

        prefix = master + "."
        for full in self.manager.sub_tables:
            if not full.startswith(prefix):
                continue
            sub = full[len(prefix):]
            fk = self.sub_fk(master, sub)
            sdf = self.manager.sub_tables[full]
            fmap = {}
            if fk and fk in sdf.columns:
                for ri, v in sdf[fk].items():
                    fmap.setdefault(str(v).strip(), []).append(ri)
            self._fk_index[(master, sub)] = fmap

    def master_idx_for_sub(self, master, sub, sub_idx):
        """Parent master row index for a sub row (via FK == PK), or None."""
        full = f"{master}.{sub}"
        sdf = self.manager.sub_tables.get(full)
        fk = self.sub_fk(master, sub)
        if sdf is None or fk not in sdf.columns or sub_idx not in sdf.index:
            return None
        fk_val = str(sdf.at[sub_idx, fk]).strip()
        return self._pk_index.get(master, {}).get(fk_val)

    def sub_rows_for_master(self, master, sub, master_idx):
        df = self.manager.tables.get(master)
        pk = self.master_pk(master)
        if df is None or pk not in df.columns or master_idx not in df.index:
            return []
        pk_val = str(df.at[master_idx, pk]).strip()
        return self._fk_index.get((master, sub), {}).get(pk_val, [])

    def count_sub_rows(self, master, master_idx, sub_name, cond_expr=""):
        """Count this master row's sub rows matching an (optional) expr."""
        rows = self.sub_rows_for_master(master, sub_name, master_idx)
        if not cond_expr.strip():
            return len(rows)
        key = ("__sub__", sub_name, cond_expr)
        se = self._sub_exprs.get(key)
        if se is None:
            try:
                se = SafeExpr(cond_expr)
            except ExprError:
                return 0
            self._sub_exprs[key] = se
        cnt = 0
        for ri in rows:
            ctx = _RowCtx(self, master, sub_name, ri)
            try:
                if se.eval(ctx):
                    cnt += 1
            except ExprError:
                return 0
        return cnt

    # ── rule evaluation ───────────────────────────────────────────────────────

    def _eval_cond(self, cond, ctx):
        agg = cond.get("agg")
        if agg:
            if ctx.scope:
                return False          # agg only valid on master scope
            field = agg.get("field", "")
            sub = agg.get("sub", "")
            rows = self.sub_rows_for_master(ctx.master, sub, ctx.row_idx)
            if field:
                op = agg.get("op", "eq")
                val, val2 = agg.get("value", ""), agg.get("value2", "")
                cnt = 0
                for ri in rows:
                    sctx = _RowCtx(self, ctx.master, sub, ri)
                    if eval_op(sctx.get(field), op, val, val2):
                        cnt += 1
            else:
                cnt = len(rows)
            cop = agg.get("count_op", "ge")
            n = _to_num(agg.get("count", 1)) or 0
            if cop in ("eq", "ne"):
                return eval_op(cnt, cop, n)
            return _cmp_num(cnt, n, cop)
        return eval_op(ctx.get(cond.get("field", "")), cond.get("op", "eq"),
                       cond.get("value", ""), cond.get("value2", ""))

    def _rule_violates(self, rule, ctx):
        """True if the rule is violated for this row (when holds, then fails)."""
        if rule.get("mode") == "expr":
            se = self._compiled.get(rule["id"])
            if se is None:
                return False          # uncompilable rule never fires
            try:
                return not bool(se.eval(ctx))
            except ExprError:
                return False
        conds = rule.get("when", {}).get("conds", [])
        if conds:
            logic = rule.get("when", {}).get("logic", "and")
            hits = (self._eval_cond(c, ctx) for c in conds)
            triggered = any(hits) if logic == "or" else all(hits)
            if not triggered:
                return False
        thens = rule.get("then", [])
        if not thens:
            return False
        return not all(self._eval_cond(c, ctx) for c in thens)

    def _mark_cols(self, rule, df):
        cols = [c for c in rule.get("mark", []) if c in df.columns]
        if cols:
            return cols
        cols = []
        for c in rule.get("then", []):
            f = c.get("field", "")
            if f and not f.startswith("master.") and not c.get("agg") \
                    and f in df.columns:
                cols.append(f)
        if cols:
            return cols
        return [df.columns[0]] if len(df.columns) else []

    def _rule_sheet(self, master, rule):
        return f"{master}.{rule['scope']}" if rule.get("scope") else master

    def _sheet_df(self, sheet):
        if sheet in self.manager.tables:
            return self.manager.tables[sheet]
        return self.manager.sub_tables.get(sheet)

    # ── violation bookkeeping ─────────────────────────────────────────────────

    def _clear_row(self, sheet, row_idx, rule_ids=None):
        """Remove violations for a row (only the given rules, or all)."""
        key = (sheet, row_idx)
        cells = self._row_cells.get(key)
        if not cells:
            return
        dead_cells = []
        for cell in cells:
            rules = self.violations.get(cell)
            if rules is None:
                dead_cells.append(cell)
                continue
            if rule_ids is None:
                rules.clear()
            else:
                rules -= rule_ids
            if not rules:
                self.violations.pop(cell, None)
                dead_cells.append(cell)
        for cell in dead_cells:
            cells.discard(cell)
        if not cells:
            self._row_cells.pop(key, None)

    def _add_violation(self, sheet, row_idx, cols, rule_id):
        for col in cols:
            cell = (sheet, row_idx, col)
            self.violations.setdefault(cell, set()).add(rule_id)
            self._row_cells.setdefault((sheet, row_idx), set()).add(cell)

    def _validate_rule_row(self, master, rule, row_idx):
        ctx = _RowCtx(self, master, rule.get("scope", ""), row_idx)
        if self._rule_violates(rule, ctx):
            sheet = self._rule_sheet(master, rule)
            df = self._sheet_df(sheet)
            if df is None:
                return
            cols = self._mark_cols(rule, df)
            # 欄位綁定判定「與此列無關」的欄位不標記——不強迫使用者
            # 補一個被隱藏/鎖定的欄位；全部無關則整筆違規不記
            binding = getattr(self.manager, "binding", None)
            if binding is not None:
                cols = [c for c in cols
                        if binding.is_relevant(sheet, row_idx, c)]
            if cols:
                self._add_violation(sheet, row_idx, cols, rule["id"])

    # ── full / per-table / per-row validation ─────────────────────────────────

    def validate_all(self):
        self.violations = {}
        self._row_cells = {}
        for master in self.manager.tables:
            self._validate_table(master, clear=False)

    def validate_table(self, master):
        """Full revalidation of one master table + its sub tables."""
        self._validate_table(master, clear=True)

    def _validate_table(self, master, clear):
        if clear:
            for sheet in [master] + [s for s in self.manager.sub_tables
                                     if s.startswith(master + ".")]:
                for key in [k for k in self._row_cells if k[0] == sheet]:
                    self._clear_row(*key)
        rules = [r for r in self.rules.get(master, []) if r.get("enabled")]
        if not rules:
            return
        self._build_indexes(master)
        for rule in rules:
            sheet = self._rule_sheet(master, rule)
            df = self._sheet_df(sheet)
            if df is None:
                continue
            for row_idx in df.index:
                self._validate_rule_row(master, rule, row_idx)

    def on_cell_edited(self, sheet, row_idx, col):
        """Incremental revalidation after one cell change."""
        if sheet in self.manager.tables:
            master, sub = sheet, None
        elif "." in sheet:
            master, sub = sheet.split(".", 1)
        else:
            return
        if master not in self.rules or not any(
                r.get("enabled") for r in self.rules.get(master, [])):
            return

        # editing the PK/FK re-links rows → cheap enough to redo the table
        pk = self.master_pk(master)
        if (sub is None and col == pk) or (sub is not None
                                           and col == self.sub_fk(master, sub)):
            self.validate_table(master)
            return
        if master not in self._pk_index:
            self._build_indexes(master)

        rules = [r for r in self.rules.get(master, []) if r.get("enabled")]
        if sub is None:
            # master row: master-scope rules + sub rules that read master.*
            self.revalidate_row(master, "", row_idx,
                                [r for r in rules if not r.get("scope")])
            for r in rules:
                scope = r.get("scope")
                if scope and self._rule_reads_master(r):
                    for si in self.sub_rows_for_master(master, scope, row_idx):
                        self.revalidate_row(master, scope, si, [r])
        else:
            sub_rules = [r for r in rules if r.get("scope") == sub]
            self.revalidate_row(master, sub, row_idx, sub_rules)
            # parent master row: master-scope rules may aggregate this sub
            midx = self.master_idx_for_sub(master, sub, row_idx)
            if midx is not None:
                m_rules = [r for r in rules if not r.get("scope")]
                if m_rules:
                    self.revalidate_row(master, "", midx, m_rules)

    def revalidate_row(self, master, scope, row_idx, rules):
        sheet = f"{master}.{scope}" if scope else master
        self._clear_row(sheet, row_idx, {r["id"] for r in rules})
        for rule in rules:
            self._validate_rule_row(master, rule, row_idx)

    @staticmethod
    def _rule_reads_master(rule):
        if rule.get("mode") == "expr":
            return "master." in rule.get("expr", "")
        for c in (rule.get("when", {}).get("conds", []) + rule.get("then", [])):
            if str(c.get("field", "")).startswith("master."):
                return True
        return False

    # ── queries for the UI ────────────────────────────────────────────────────

    def cell_rules(self, sheet, row_idx, col):
        """Rules violated at a cell, errors first."""
        ids = self.violations.get((sheet, row_idx, col))
        if not ids:
            return []
        rules = [self.rule_by_id[i][1] for i in ids if i in self.rule_by_id]
        return sorted(rules, key=lambda r: 0 if r["severity"] == "error" else 1)

    def cell_color(self, sheet, row_idx, col):
        rules = self.cell_rules(sheet, row_idx, col)
        return rules[0]["color"] if rules else None

    def row_has_violation(self, sheet, row_idx):
        return (sheet, row_idx) in self._row_cells

    def sheet_violation_rows(self, sheet):
        return {k[1] for k in self._row_cells if k[0] == sheet}

    def record_has_violation(self, master, master_idx):
        """True if the master row or any of its sub rows has a violation."""
        return self.record_violation_severity(master, master_idx) is not None

    def record_violation_severity(self, master, master_idx):
        """Worst severity for a master row incl. its sub rows:
        'error' > 'warn' > None."""
        worst = None

        def _row_worst(sheet, row_idx):
            nonlocal worst
            for cell in self._row_cells.get((sheet, row_idx), ()):
                for rid in self.violations.get(cell, ()):
                    info = self.rule_by_id.get(rid)
                    if info:
                        if info[1]["severity"] == "error":
                            return True     # can't get worse
                        worst = "warn"
            return False

        if _row_worst(master, master_idx):
            return "error"
        prefix = master + "."
        for full in self.manager.sub_tables:
            if not full.startswith(prefix):
                continue
            sub = full[len(prefix):]
            for si in self.sub_rows_for_master(master, sub, master_idx):
                if _row_worst(full, si):
                    return "error"
        return worst

    def has_errors(self):
        for ids in self.violations.values():
            for i in ids:
                info = self.rule_by_id.get(i)
                if info and info[1]["severity"] == "error":
                    return True
        return False

    def summary(self):
        """Flatten violations for the save-gate dialog.

        Returns [{severity, rule, master, sheet, is_sub, row_idx, cols, pk_val}]
        sorted errors first, then by rule name.
        """
        by_rule_row = {}   # (rule_id, sheet, row_idx) -> [cols]
        for (sheet, row_idx, col), ids in self.violations.items():
            for i in ids:
                by_rule_row.setdefault((i, sheet, row_idx), []).append(col)
        out = []
        for (rid, sheet, row_idx), cols in by_rule_row.items():
            info = self.rule_by_id.get(rid)
            if not info:
                continue
            master, rule = info
            is_sub = sheet != master
            pk_val = ""
            if is_sub:
                midx = self.master_idx_for_sub(master, sheet.split(".", 1)[1],
                                               row_idx)
            else:
                midx = row_idx
            mdf = self.manager.tables.get(master)
            pk = self.master_pk(master)
            if mdf is not None and midx is not None and pk in mdf.columns \
                    and midx in mdf.index:
                pk_val = str(mdf.at[midx, pk])
            out.append({
                "severity": rule["severity"], "rule": rule, "master": master,
                "sheet": sheet, "is_sub": is_sub, "row_idx": row_idx,
                "cols": sorted(cols), "pk_val": pk_val,
            })
        out.sort(key=lambda d: (0 if d["severity"] == "error" else 1,
                                d["rule"]["name"], d["sheet"], str(d["row_idx"])))
        return out

    def count_by_severity(self):
        errs = warns = 0
        for item in self.summary():
            if item["severity"] == "error":
                errs += 1
            else:
                warns += 1
        return errs, warns

    # ── ad-hoc rule test (rules-editor「測試」button) ──────────────────────────

    def test_rule(self, master, rule):
        """Run one rule against current data without touching the violation map.
        Returns (count, sample_pk_list, error_message_or_None)."""
        rule = normalize_rule(dict(rule))
        se = None
        if rule.get("mode") == "expr":
            try:
                se = SafeExpr(rule.get("expr", ""))
            except ExprError as e:
                return 0, [], str(e)
        self._build_indexes(master)
        old = self._compiled.get(rule["id"])
        if se is not None:
            self._compiled[rule["id"]] = se
        sheet = self._rule_sheet(master, rule)
        df = self._sheet_df(sheet)
        cnt, samples = 0, []
        if df is not None:
            for row_idx in df.index:
                ctx = _RowCtx(self, master, rule.get("scope", ""), row_idx)
                try:
                    bad = self._rule_violates(rule, ctx)
                except ExprError as e:
                    self._restore_compiled(rule["id"], old)
                    return 0, [], str(e)
                if bad:
                    cnt += 1
                    if len(samples) < 10:
                        samples.append(self._row_label(master, rule, row_idx))
        self._restore_compiled(rule["id"], old)
        return cnt, samples, None

    def _restore_compiled(self, rid, old):
        if old is None:
            self._compiled.pop(rid, None)
        else:
            self._compiled[rid] = old

    def _row_label(self, master, rule, row_idx):
        scope = rule.get("scope", "")
        if scope:
            midx = self.master_idx_for_sub(master, scope, row_idx)
        else:
            midx = row_idx
        mdf = self.manager.tables.get(master)
        pk = self.master_pk(master)
        if mdf is not None and midx is not None and pk in mdf.columns \
                and midx in mdf.index:
            return str(mdf.at[midx, pk])
        return f"row {row_idx}"
