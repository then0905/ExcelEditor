"""Conditional field visibility engine (Qt-free).

A table (master or sub) may declare ONE binding in its master's config:

    "field_bindings": {
      "<scope>": {                    # "" = the master table, else sub name
        "enabled": true,
        "driver": "SkillComponent",
        "groups": { "<driver value>": ["FieldA", "FieldB", ...], ... }
      }
    }

Semantics (user-approved design, see docs/superpowers/specs/
2026-07-12-field-binding-design.md):

- Fields NOT mentioned in any group are shared fields → always relevant.
- The driver column, the sub table's FK and the master's PK are always relevant.
- A row whose driver value has no group entry shows everything (never
  mis-hide data for values that simply haven't been configured yet).
- Binding only controls visibility/editability — values are never mutated
  (decision "B": hidden fields keep whatever value they had).
"""


class FieldBindingEngine:
    def __init__(self, manager):
        self.manager = manager

    # ── config access ─────────────────────────────────────────────────────────

    def binding_for(self, master, scope):
        """The binding dict for a table scope, or None if absent/disabled."""
        b = (self.manager.config.get(master, {})
             .get("field_bindings", {}).get(scope))
        if not isinstance(b, dict) or not b.get("enabled", True):
            return None
        driver = b.get("driver", "")
        if not driver or not isinstance(b.get("groups"), dict):
            return None
        return b

    def _always_relevant(self, master, scope, binding, columns):
        cfg = self.manager.config.get(master, {})
        keep = {binding.get("driver", "")}
        if scope:
            keep.add(cfg.get("sub_tables", {}).get(scope, {})
                     .get("foreign_key") or cfg.get("primary_key", ""))
        else:
            keep.add(cfg.get("primary_key", ""))
        mentioned = set()
        for fields in binding.get("groups", {}).values():
            if isinstance(fields, list):
                mentioned.update(str(f) for f in fields)
        keep.update(c for c in columns if c not in mentioned)   # shared fields
        return keep

    def _sheet_df(self, master, scope):
        if scope:
            return self.manager.sub_tables.get(f"{master}.{scope}")
        return self.manager.tables.get(master)

    # ── queries ───────────────────────────────────────────────────────────────

    def relevant_fields(self, master, scope, row_idx):
        """Set of relevant column names for one row, or None = everything
        (no binding / row missing / driver value not configured)."""
        binding = self.binding_for(master, scope)
        if binding is None:
            return None
        df = self._sheet_df(master, scope)
        driver = binding["driver"]
        if df is None or driver not in df.columns or row_idx not in df.index:
            return None
        val = str(df.at[row_idx, driver]).strip()
        group = binding.get("groups", {}).get(val)
        if not isinstance(group, list):
            return None                       # unconfigured value → show all
        cols = list(df.columns)
        rel = self._always_relevant(master, scope, binding, cols)
        rel.update(str(f) for f in group)
        return rel

    def is_relevant(self, sheet, row_idx, col):
        """Cell-level check; sheet is a master name or 'Master.Sub'."""
        if sheet in self.manager.tables:
            master, scope = sheet, ""
        elif "." in sheet:
            master, scope = sheet.split(".", 1)
        else:
            return True
        rel = self.relevant_fields(master, scope, row_idx)
        return True if rel is None else col in rel

    def visible_columns(self, master, scope, row_idxs):
        """Union of relevant columns over the given rows (for grid column
        hiding), or None = show every column."""
        binding = self.binding_for(master, scope)
        if binding is None:
            return None
        df = self._sheet_df(master, scope)
        if df is None:
            return None
        visible = set()
        for ri in row_idxs:
            rel = self.relevant_fields(master, scope, ri)
            if rel is None:
                return None                   # one unconfigured row → show all
            visible |= rel
        if not row_idxs:
            return None                       # empty view → nothing to hide
        return visible

    def driver_of(self, master, scope):
        b = self.binding_for(master, scope)
        return b["driver"] if b else None
