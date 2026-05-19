"""GET /external-functions — 外部 Add-In 関数レジストリの公開.

フロントエンドの「外部関数」タブが、当ツールが知っている関数定義 (Bloomberg
BDH/BDP/BDS 等) と、当該ジョブで実際に使われている関数の集計を取得するための
エンドポイント.
"""

from __future__ import annotations

import logging
from collections import Counter
from typing import Any

from fastapi import APIRouter, Depends, HTTPException

from backend.dependencies import get_storage
from backend.storage import JobNotFoundError, Storage
from core.external_functions import get_function, list_functions

logger = logging.getLogger(__name__)
router = APIRouter()


@router.get("/external-functions")
async def get_external_functions() -> dict[str, Any]:
    """全ベンダーの登録済み関数一覧を返す.

    ジョブ非依存. 設計書ページの「外部関数」タブが定義データを取得する用途.
    """
    return {
        "functions": [fn.model_dump() for fn in list_functions()],
        "vendors": sorted({fn.vendor for fn in list_functions()}),
    }


@router.get("/external-functions/used/{job_id}")
async def get_external_functions_used(
    job_id: str,
    storage: Storage = Depends(get_storage),
) -> dict[str, Any]:
    """当該ジョブの Workbook で使われている外部関数の使用状況を返す.

    Returns:
        {
          "items": [
            {"name": "BDH", "vendor": "Bloomberg", "count": 47,
             "locations": [{"sheet": "Calc", "coord": "H2", "formula": "..."}, ...]}
          ],
          "total_kinds": int, "total_uses": int
        }
    """
    try:
        wb = storage.load_workbook(job_id)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=f"invalid job_id: {e}") from e
    except JobNotFoundError as e:
        raise HTTPException(status_code=404, detail="job not found") from e
    except FileNotFoundError as e:
        raise HTTPException(
            status_code=409,
            detail="workbook not extracted yet; call /extract first",
        ) from e

    counts: Counter[str] = Counter()
    locations: dict[str, list[dict[str, str]]] = {}
    for sheet in wb.sheets:
        for f in sheet.formulas:
            for fn_name in f.external_functions:
                counts[fn_name] += 1
                # 全使用箇所を保持 (フロントで一覧表示)
                locations.setdefault(fn_name, []).append(
                    {
                        "sheet": sheet.name,
                        "coord": f.coord if "!" in f.coord else f"{sheet.name}!{f.coord}",
                        "formula": f.formula,
                    }
                )

    items = []
    for fn_name, cnt in counts.most_common():
        fn = get_function(fn_name)
        items.append(
            {
                "name": fn_name,
                "vendor": fn.vendor if fn else "?",
                "short": fn.short if fn else "",
                "count": cnt,
                "locations": locations.get(fn_name, []),
                "registered": fn is not None,
            }
        )

    return {
        "items": items,
        "total_kinds": len(items),
        "total_uses": sum(counts.values()),
    }
