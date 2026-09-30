# -*- coding: utf-8 -*-
"""Vẽ biểu đồ Excel gốc bằng XlsxWriter từ cấu hình trong bieu_do.py."""

PHONG = "Times New Roman"
MAU_CHU = "#262522"
MAU_LUOI = "#DEDDD8"
MAU_TRUC = "#8C8B86"


def _font(size=10, bold=False, color=MAU_CHU):
    return {"name": PHONG, "size": size, "bold": bold, "color": color}


def ghi_bang(ws, hang0, cot0, hinh, dd_tieu_de, dd_o, dd_so):
    """Ghi bảng dữ liệu của một hình; trả về (hàng đầu, hàng cuối) của phần dữ liệu."""
    for j, c in enumerate(hinh["cot"]):
        ws.write(hang0, cot0 + j, c, dd_tieu_de)
    for i, dong in enumerate(hinh["dong"]):
        for j, v in enumerate(dong):
            fmt = hinh["dinh_dang"][j]
            if isinstance(v, (int, float)) and fmt:
                ws.write_number(hang0 + 1 + i, cot0 + j, v, dd_so[fmt])
            else:
                ws.write(hang0 + 1 + i, cot0 + j, v, dd_o)
    return hang0 + 1, hang0 + len(hinh["dong"])


def _ref(sheet, r, c):
    from xlsxwriter.utility import xl_rowcol_to_cell
    return f"='{sheet}'!{xl_rowcol_to_cell(r, c, row_abs=True, col_abs=True)}"


def _range(sheet, r1, c1, r2, c2):
    from xlsxwriter.utility import xl_range_abs
    return f"='{sheet}'!{xl_range_abs(r1, c1, r2, c2)}"


def ve(wb, sheet, hinh, hang_tieu_de, cot0, r1, r2, tieu_de=None):
    """Tạo đối tượng biểu đồ cho một hình. hang_tieu_de: hàng chứa tên cột."""
    cfg = hinh["bieu_do"]
    loai = cfg["loai"]
    cat = _range(sheet, r1, cot0, r2, cot0)

    if loai == "doughnut":
        ch = wb.add_chart({"type": "doughnut"})
        ch.add_series({
            "name": _ref(sheet, hang_tieu_de, cot0 + 1),
            "categories": cat,
            "values": _range(sheet, r1, cot0 + 1, r2, cot0 + 1),
            "points": [{"fill": {"color": m}, "border": {"color": "#FFFFFF", "width": 1.5}} for m in cfg["mau_diem"]],
            "data_labels": {"percentage": True, "value": True, "separator": "\n", "font": _font(11, True, "#FFFFFF")},
        })
        ch.set_hole_size(cfg.get("lo", 55))
        ch.set_legend({"position": "right", "font": _font(10)})
        _chung(ch, tieu_de)
        return ch

    subtype = "stacked" if cfg.get("xep_chong") else None
    loai_chinh = {"type": loai}
    if subtype:
        loai_chinh["subtype"] = subtype
    ch = wb.add_chart(loai_chinh)
    phu = None
    for s in cfg["chuoi"]:
        kieu = s.get("kieu", loai)
        target = ch
        if kieu != loai:
            if phu is None:
                phu = wb.add_chart({"type": kieu})
            target = phu
        opt = {
            "name": _ref(sheet, hang_tieu_de, cot0 + s["cot"]),
            "categories": cat,
            "values": _range(sheet, r1, cot0 + s["cot"], r2, cot0 + s["cot"]),
        }
        if kieu in ("column", "bar"):
            opt["fill"] = {"color": s["mau"]}
            opt["border"] = {"color": "#FFFFFF", "width": 0.75}
            opt["gap"] = cfg.get("khoang_cach", 60)
            if subtype:
                opt["overlap"] = 100
            if s.get("mau_diem"):
                opt["points"] = [{"fill": {"color": m}, "border": {"color": "#FFFFFF", "width": 0.75}}
                                 for m in s["mau_diem"]]
        else:
            line = {"color": s["mau"], "width": s.get("dam", 2)}
            if s.get("gach"):
                line["dash_type"] = "dash"
            if s.get("cham"):
                line["dash_type"] = "round_dot"
            if s.get("duong", True) is False:
                opt["line"] = {"none": True}
                opt["marker"] = {"type": "circle", "size": 8, "fill": {"color": s["mau"]},
                                 "border": {"color": "#FFFFFF", "width": 1}}
            else:
                opt["line"] = line
                opt["smooth"] = False
                opt["marker"] = {"type": "circle", "size": 6, "fill": {"color": s["mau"]},
                                 "border": {"color": s["mau"]}}
        if s.get("nhan"):
            dl = {"value": True, "font": _font(9, False, MAU_CHU)}
            if s.get("vi_tri_nhan"):
                dl["position"] = s["vi_tri_nhan"]
            elif kieu in ("column", "bar") and not subtype:
                dl["position"] = "outside_end"
            elif subtype:
                dl["position"] = "center"
                dl["font"] = _font(9, True, "#FFFFFF")
            if s.get("an_nhan_0"):
                dl["custom"] = [{"delete": True} if (d[s["cot"]] or 0) == 0 else None for d in hinh["dong"]]
            fmt = s.get("dd_nhan") or hinh["dinh_dang"][s["cot"]]
            if fmt:
                dl["num_format"] = fmt
            opt["data_labels"] = dl
        if s.get("truc_phu"):
            opt["y2_axis"] = True
        target.add_series(opt)

    truc_gia_tri = {
        "num_font": _font(10), "name_font": _font(10, True),
        "line": {"color": MAU_TRUC, "width": 0.75},
        "major_gridlines": {"visible": True, "line": {"color": MAU_LUOI, "width": 0.5}},
        "major_tick_mark": "none", "minor_tick_mark": "none",
    }
    if cfg.get("max_y"):
        truc_gia_tri["max"] = cfg["max_y"]
    if cfg.get("buoc_y"):
        truc_gia_tri["major_unit"] = cfg["buoc_y"]
    if cfg.get("truc_y"):
        truc_gia_tri["name"] = cfg["truc_y"]
    truc_danh_muc = {"num_font": _font(10), "line": {"color": MAU_TRUC, "width": 0.75},
                     "major_tick_mark": "none", "minor_tick_mark": "none"}
    if cfg.get("truc_x"):
        truc_danh_muc["name"] = cfg["truc_x"]
        truc_danh_muc["name_font"] = _font(10, True)

    if loai == "bar":
        if cfg.get("dao_truc"):
            truc_danh_muc["reverse"] = True
            truc_danh_muc["crossing"] = "max"
        ch.set_y_axis(truc_danh_muc)
        truc_gia_tri["min"] = 0
        ch.set_x_axis(truc_gia_tri)
    else:
        ch.set_x_axis(truc_danh_muc)
        truc_gia_tri["min"] = 0
        ch.set_y_axis(truc_gia_tri)

    if phu is not None:
        if cfg.get("truc_y2"):
            y2 = {"name": cfg["truc_y2"], "name_font": _font(10, True), "num_font": _font(10),
                  "line": {"color": MAU_TRUC, "width": 0.75}, "major_tick_mark": "none",
                  "minor_tick_mark": "none", "min": 0, "major_gridlines": {"visible": False}}
            if cfg.get("dd_y2"):
                y2["num_format"] = cfg["dd_y2"]
            if cfg.get("max_y2"):
                y2["max"] = cfg["max_y2"]
            phu.set_y2_axis(y2)
        ch.combine(phu)

    if cfg.get("an_chu_giai"):
        ch.set_legend({"none": True})
    else:
        ch.set_legend({"position": "bottom", "font": _font(10)})
    _chung(ch, tieu_de)
    return ch


def _chung(ch, tieu_de):
    if tieu_de:
        ch.set_title({"name": tieu_de, "name_font": _font(12, True)})
    else:
        ch.set_title({"none": True})
    ch.set_chartarea({"border": {"none": True}, "fill": {"color": "#FFFFFF"}})
    ch.set_plotarea({"border": {"none": True}})
