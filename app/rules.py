"""Catálogo de governança transcrito das referências fornecidas.

Listas corporativas são deliberadamente externas: a ausência de uma lista nunca
transforma um valor observado no YAML em um valor oficial.
"""
from __future__ import annotations

import copy
import re
from datetime import date
from typing import Any

BOOL = ["true", "false"]
CLASSIFICATIONS = ["publica", "interna", "confidencial", "restrita", "sigilosa"]
LEGAL = ["consentimento", "execucao_de_contrato", "obrigacao_legal", "legitimo_interesse",
         "protecao_ao_credito", "protecao_da_vida", "tutela_da_saude",
         "exercicio_regular_de_direitos", "politicas_publicas", "estudo_e_pesquisa", "nao_aplicavel"]
LAYERS = ["bronze", "silver", "gold", "data_product"]
RISK = ["baixo", "medio", "alto", "critico"]
REFERENCE_ALIASES = {
    "nm_dom": ("nm_dom", "dominios", "domains"),
    "nm_subdom": ("nm_subdom", "subdominios", "subdomains"),
    "nm_sist_orig": ("nm_sist_orig", "sistemas_origem", "source_systems"),
    "ct_area_neg": ("ct_area_neg", "areas_negocio", "business_areas"),
}


def _entry(description: str, values=None, *, format=None, external=None):
    result = {"description": description, "category": "Não identificada", "values": values}
    if format:
        result["format"] = format
    if external:
        result["external"] = external
    return result


COMMON = {
    "nm_environment": _entry("Ambiente em que o objeto está publicado.", ["dev", "prd"]),
    "dt_aprv_gov": _entry("Data da aprovação pela Governança de Dados.", format="YYYY-MM-DD"),
    "nm_cust": _entry("Custodiante responsável pelo objeto.", format="email"),
    "nm_dom": _entry("Domínio oficial de dados.", external="nm_dom"),
    "fl_reg_apl": _entry("Indica aplicação de requisitos regulatórios.", BOOL),
    "ct_risco": _entry("Categoria de risco do objeto.", RISK),
    "nm_class_inf": _entry("Classificação da informação.", CLASSIFICATIONS),
    "ct_area_neg": _entry("Área de negócio oficial associada ao objeto.", external="ct_area_neg"),
}


def _catalog():
    catalog = copy.deepcopy(COMMON)
    catalog.update({
        "grp_owner": _entry("Grupo de responsáveis pelo catálogo.", format="grp-dbx-data-{dominio}-owners"),
        "nm_environment": _entry("Ambiente em que o catálogo está publicado.", ["lab", "dev", "prd"]),
        "fl_exempt": _entry("Indica a isenção prevista no padrão de governança.", BOOL),
        "nm_purpose": _entry("Finalidade de governança do catálogo.", ["governed", "not_governed"]),
        "grp_own_stwd": _entry("Grupo de Data Stewards do domínio.", format="grp-dbx-data-{dominio}-stewards"),
        "nm_cust": _entry("Grupo ou e-mail corporativo do custodiante.", format="group_or_email"),
        "fl_ropa": _entry("Indica registro das operações de tratamento.", BOOL),
        "fl_metadata_ai_generate": _entry("Indica metadados gerados por inteligência artificial.", BOOL),
    })
    schema = copy.deepcopy(COMMON)
    schema.update({
        "grp_own_stwd": _entry("Grupo de Data Stewards do domínio.", format="grp-dbx-data-{dominio}-stewards"),
        "nm_cust": _entry("Grupo ou e-mail corporativo do custodiante.", format="group_or_email"),
        "nm_cam": _entry("Camada de organização dos dados.", LAYERS),
        "nm_lgl_lgpd": _entry("Base legal do tratamento de dados pessoais.", LEGAL),
        "fl_ropa": _entry("Indica registro das operações de tratamento.", BOOL),
    })
    table = copy.deepcopy(COMMON)
    table.update({
        "nm_own_stwd": _entry("E-mail do responsável de negócio ou Data Steward.", format="email"),
        "nm_subdom": _entry("Subdomínio oficial de dados.", external="nm_subdom"),
        "nm_cam": _entry("Camada de organização dos dados.", LAYERS),
        "nm_sist_orig": _entry("Sistema de origem oficial ou múltiplas origens.", external="nm_sist_orig"),
        "nm_lgl_lgpd": _entry("Base legal do tratamento de dados pessoais.", LEGAL),
        "fl_pii": _entry("Indica presença de dados pessoais.", BOOL),
        "fl_sens": _entry("Indica presença de dados pessoais sensíveis.", BOOL),
        "nm_retencao": _entry("Prazo de retenção dos dados.", ["1_ano", "5_anos", "7_anos", "10_anos", "indeterminada"]),
        "fl_mask": _entry("Indica mascaramento em pelo menos uma coluna.", BOOL),
        "ct_qld_status": _entry("Situação da avaliação de qualidade.", ["nao_avaliado", "em_avaliacao", "aprovado", "reprovado"]),
        "fl_apt_cons": _entry("Indica aptidão para consumo após qualidade aprovada e aprovação da governança.", BOOL),
        "ct_cons_exec": _entry("Indicador de consumo executivo.", BOOL),
        "fl_legado": _entry("Indica objeto legado.", BOOL),
        "ct_prod_neg": _entry("Produto de negócio escrito em snake_case.", format="snake_case"),
        "fl_metadata_ai_generate": _entry("Indica metadados gerados por inteligência artificial.", BOOL),
    })
    view = copy.deepcopy(table)
    del view["nm_sist_orig"]  # Diferença expressa entre as tabelas oficiais de Table e View.
    column = {
        "fl_reg_apl": _entry("Indica aplicação de requisitos regulatórios à coluna.", BOOL),
        "fl_atr_pii": _entry("Indica dado pessoal na coluna.", BOOL),
        "nm_atr_pii": _entry("Tipo de dado pessoal da coluna.", ["cpf", "rg", "nome", "email", "telefone", "endereco", "data_nascimento", "dado_financeiro", "geolocalizacao", "biometria", "saude", "outro", "nao_aplicavel"]),
        "fl_atr_sens": _entry("Indica dado pessoal sensível na coluna.", BOOL),
        "fl_atr_mask": _entry("Indica mascaramento na coluna.", BOOL),
        "nm_atr_class_inf": _entry("Classificação da informação contida na coluna.", CLASSIFICATIONS),
        "fl_metadata_ai_generate": _entry("Indica metadados gerados por inteligência artificial.", BOOL),
    }
    return {"Catalog": catalog, "Schema": schema, "Table": table, "View": view, "Column": column}


def get_catalog(references: dict | None = None) -> dict:
    result = _catalog()
    references = {} if references is None else references
    if not isinstance(references, dict):
        raise ValueError("Referências devem ser um mapa de listas oficiais.")
    for level in result.values():
        for rule in level.values():
            key = rule.get("external")
            if not key:
                continue
            for alias in REFERENCE_ALIASES[key]:
                if alias in references:
                    values = references[alias]
                    if not isinstance(values, list) or any(not isinstance(x, str) or not x.strip() for x in values):
                        raise ValueError(f"Lista de referência {key} deve conter apenas textos não vazios.")
                    rule["values"] = list(dict.fromkeys(values))
                    if key == "nm_sist_orig" and "multiplas_origens" not in rule["values"]:
                        rule["values"].append("multiplas_origens")
                    break
    return result


def validate_tag(level: str, tag: str, value: Any, references=None, *, catalog=None) -> list[dict]:
    """Valida um valor sem converter booleanos nativos nem criar padrões."""
    rule = (catalog if catalog is not None else get_catalog(references)).get(level, {}).get(tag)
    def issue(code, severity, message):
        return [{"field": f"tags.{tag}", "code": code, "severity": severity, "message": message}]
    if isinstance(value, bool):
        return issue("native_boolean", "error", 'Booleanos de tags devem ser textos "true" ou "false".')
    if not isinstance(value, str):
        return issue("invalid_type", "error", "O valor da tag deve ser texto.")
    if rule is None:
        return issue("unknown_tag", "pending", "Tag não documentada neste nível; preservada para revisão.")
    if not value.strip():
        return issue("missing_value", "pending", "Tag sem valor; requer preenchimento ou decisão do responsável.")
    if rule.get("external") and rule["values"] is None:
        if tag == "nm_sist_orig" and value == "multiplas_origens":
            return []
        return issue("external_list_missing", "pending", "Lista oficial não carregada")
    if rule["values"] is not None and value not in rule["values"]:
        return issue("not_allowed", "error", "Valor fora da lista oficial de valores permitidos.")
    fmt = rule.get("format")
    valid = True
    if fmt == "YYYY-MM-DD":
        valid = bool(re.fullmatch(r"\d{4}-\d{2}-\d{2}", value))
        if valid:
            try:
                date.fromisoformat(value)
            except ValueError:
                valid = False
    elif fmt == "email":
        valid = bool(re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", value))
    elif fmt == "group_or_email":
        valid = bool(re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", value) or re.fullmatch(r"grp-[a-zA-Z0-9][a-zA-Z0-9_-]*", value))
    elif fmt == "snake_case":
        valid = bool(re.fullmatch(r"[a-z][a-z0-9]*(?:_[a-z0-9]+)*", value))
    elif fmt in ("grp-dbx-data-{dominio}-owners", "grp-dbx-data-{dominio}-stewards"):
        suffix = "owners" if fmt.endswith("owners") else "stewards"
        valid = bool(re.fullmatch(r"grp-dbx-data-[a-z0-9][a-z0-9_-]*-" + suffix, value))
    if not valid:
        return issue("invalid_format", "error", f"Formato esperado: {fmt}.")
    return []


def validate_nodes(nodes: list[dict], references: dict | None = None) -> list[dict]:
    catalog = get_catalog(references)  # Rejeita referências malformadas mesmo sem tags.
    issues: list[dict] = []
    if any(catalog.get(node["level"], {}).get(tag, {}).get("category") == "Não identificada"
           for node in nodes for tag in node.get("tags", {})):
        issues.append({"id": "", "field": "catalog.category", "code": "category_not_identified",
                       "severity": "pending", "scope": "governance",
                       "message": "Categorias de tags não identificadas no padrão: revisão da Governança necessária; categorias não foram inferidas."})
    by_path = {(node["level"], node["path"]): node for node in nodes}
    columns: dict[str, list[dict]] = {}
    for node in nodes:
        if node["level"] == "Column":
            columns.setdefault(node.get("object_id", ""), []).append(node)
        for tag, value in node.get("tags", {}).items():
            if tag in node.get("_native_tag_types", {}) and isinstance(value, str):
                issues.append({"id": node["id"], "field": f"tags.{tag}", "code": "native_tag_type", "severity": "error",
                               "message": "Tag armazenada com tipo YAML nativo; confirme explicitamente o valor como texto."})
            else:
                issues.extend({"id": node["id"], **i} for i in validate_tag(node["level"], tag, value, references, catalog=catalog))
    for node in nodes:
        if node["level"] not in ("Table", "View"):
            continue
        tags = node.get("tags", {})
        children = columns.get(node["id"], [])
        def add(field, code, severity, message):
            issues.append({"id": node["id"], "field": field, "code": code, "severity": severity, "message": message})
        if any(c.get("tags", {}).get("fl_atr_mask") == "true" for c in children) and tags.get("fl_mask") != "true":
            add("tags.fl_mask", "mask_inconsistent", "error", 'Há coluna mascarada; fl_mask do objeto deve ser "true".')
        if tags.get("fl_apt_cons") == "true":
            if tags.get("ct_qld_status") != "aprovado":
                add("tags.fl_apt_cons", "consumption_quality", "error", "Aptidão para consumo exige ct_qld_status aprovado.")
            approval = tags.get("dt_aprv_gov")
            if not approval or any(i["severity"] == "error" for i in validate_tag(node["level"], "dt_aprv_gov", approval, references, catalog=catalog)):
                add("tags.dt_aprv_gov", "consumption_approval", "error", "Aptidão para consumo exige data válida de aprovação da governança.")
        actual = tags.get("nm_class_inf")
        classified = [c.get("tags", {}).get("nm_atr_class_inf") for c in children]
        known = [v for v in classified if v in CLASSIFICATIONS]
        if known and actual in CLASSIFICATIONS and CLASSIFICATIONS.index(actual) < max(CLASSIFICATIONS.index(v) for v in known):
            add("tags.nm_class_inf", "classification_too_low", "error", "Classificação menos restritiva que a informação de pelo menos uma coluna.")
        if known or actual is not None:
            upper = [by_path.get(("Catalog", node.get("catalog", ""))), by_path.get(("Schema", f'{node.get("catalog", "")}.{node.get("schema", "")}'))]
            safe_upper = all(u and u.get("tags", {}).get("nm_class_inf") in CLASSIFICATIONS for u in upper)
            if not children or len(known) != len(children) or actual not in CLASSIFICATIONS or not safe_upper:
                add("tags.nm_class_inf", "classification_review", "pending", "Não é possível calcular a classificação com segurança: revisar colunas e limites dos níveis superiores.")
            elif any(CLASSIFICATIONS.index(actual) > CLASSIFICATIONS.index(u["tags"]["nm_class_inf"]) for u in upper):
                add("tags.nm_class_inf", "classification_upper_limit", "pending", "A classificação excede a dos níveis superiores; o limite documentado requer revisão de negócio.")
    return issues
