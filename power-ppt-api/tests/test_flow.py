import base64
import io

import pytest
from PIL import Image as PILImage

from app.brand_engine import geometry as g
from app.brand_engine.flow import flow_pages
from app.schemas import Image, Page, Table


def _img_b64(w=800, h=400) -> str:
    im = PILImage.new("RGB", (w, h), (122, 193, 67))
    b = io.BytesIO()
    im.save(b, "PNG")
    return base64.b64encode(b.getvalue()).decode()


def _titles(specs):
    return [s["title"] for s in specs]


def _kinds(spec):
    return [p["kind"] for p in spec["placements"]]


def test_short_page_stays_single_slide():
    specs = flow_pages([Page(title="Intro", body="One short line.")])
    assert len(specs) == 1
    assert specs[0]["title"] == "Intro"


def test_long_body_spills_with_contd_titles():
    body = "\n\n".join(f"Paragraph {i}: " + ("cover detail " * 20) for i in range(20))
    specs = flow_pages([Page(title="Details", body=body)])
    assert len(specs) > 1
    assert specs[0]["title"] == "Details"
    assert all(t == "Details" + g.CONTINUATION_SUFFIX for t in _titles(specs)[1:])


def test_table_splits_by_rows_and_repeats_header():
    rows = [["Sr", "Item"]] + [[str(i), f"row {i}"] for i in range(1, 41)]
    specs = flow_pages([Page(title="Sched", tables=[Table(header=rows[0], rows=rows)])])
    assert len(specs) > 1
    total_data = 0
    for spec in specs:
        for p in spec["placements"]:
            assert p["kind"] == "table"
            tbl = p["payload"]
            assert tbl.rows[0] == ["Sr", "Item"]   # header repeated on every part
            total_data += len(tbl.rows) - 1
    assert total_data == 40                          # no rows lost


def test_image_gets_its_own_slide_after_full_body():
    body = "\n\n".join(f"Paragraph {i}: " + ("cover detail " * 20) for i in range(20))
    specs = flow_pages([Page(title="Mix", body=body, images=[Image(data=_img_b64())])])
    # the image must land on some slide, exactly once
    img_slides = [i for i, s in enumerate(specs) if "image" in _kinds(s)]
    assert len(img_slides) == 1


def test_oversized_single_paragraph_is_split():
    # one paragraph far taller than a whole content zone must still be broken up
    huge = "word " * 4000
    specs = flow_pages([Page(title="Big", body=huge)])
    assert len(specs) > 1
    assert all("body" in _kinds(s) for s in specs)


def test_table_and_image_pack_side_by_side():
    # a small table + an image should share one row (same top, different lefts),
    # each at half width, below the body.
    table = Table(header=["A", "B"], rows=[["A", "B"], ["1", "2"]])
    specs = flow_pages([Page(title="Mix", body="Intro line.",
                             tables=[table], images=[Image(data=_img_b64())])])
    assert len(specs) == 1
    pls = specs[0]["placements"]
    tbl = next(p for p in pls if p["kind"] == "table")
    img = next(p for p in pls if p["kind"] == "image")
    assert abs(tbl["top"] - img["top"]) < 1e-6          # same row
    assert tbl["left"] < img["left"]                     # side by side
    assert abs(tbl["width"] - g.HALF_WIDTH) < 1e-6       # half width each
    assert abs(img["width"] - g.HALF_WIDTH) < 1e-6
    body = next(p for p in pls if p["kind"] == "body")
    assert body["top"] < tbl["top"]                      # text above the row


def test_wide_table_stays_full_width():
    wide = Table(header=list("ABCDEF"), rows=[list("ABCDEF"), list("123456")])
    specs = flow_pages([Page(title="Wide", tables=[wide], images=[Image(data=_img_b64())])])
    pls = specs[0]["placements"]
    tbl = next(p for p in pls if p["kind"] == "table")
    assert abs(tbl["width"] - g.BODY_WIDTH) < 1e-6       # 6 cols -> not paired


def test_empty_page_yields_one_titled_slide():
    specs = flow_pages([Page(title="Section only")])
    assert len(specs) == 1
    assert specs[0]["title"] == "Section only"
    assert specs[0]["placements"] == []
