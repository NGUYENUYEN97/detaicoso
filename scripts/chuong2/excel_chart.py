# -*- coding: utf-8 -*-
"""Vẽ biểu đồ Excel gốc bằng XlsxWriter từ cấu hình trong bieu_do.py."""

PHONG = "Times New Roman"
MAU_CHU = "#262522"
MAU_LUOI = "#DEDDD8"
MAU_TRUC = "#8C8B86"


def _font(size=10, bold=False, color=MAU_CHU):
    return {"name": PHONG, "size": size, "bold": bold, "color": color}


def _do_sang(hex_mau):
    """Độ sáng tương đối theo WCAG 2.1."""
    h = hex_mau.lstrip("#")
    kq = []
    for i in (0, 2, 4):
        c = int(h[i:i + 2], 16) / 255
        kq.append(c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4)
    return 0.2126 * kq[0] + 0.7152 * kq[1] + 0.0722 * kq[2]


def mau_chu_tren(nen):
    """Chọn chữ trắng hoặc chữ sẫm, lấy màu có độ tương phản cao hơn với màu nền."""
    l = _do_sang(nen)
    tuong_phan_trang = 1.05 / (l + 0.05)
    tuong_phan_sam = (l + 0.05) / (_do_sang(MAU_CHU) + 0.05)
    return "#FFFFFF" if tuong_phan_trang >= tuong_phan_sam else MAU_CHU


# ---------------------------------------------------------------------------
# Bố cục cố định (manual layout): chú giải đặt phía trên, vùng vẽ đặt bên dưới,
# hai vùng không chồng lên nhau trong Word, Excel và LibreOffice.
# Đơn vị tính gần đúng theo cỡ chữ 10 pt Times New Roman.
# ---------------------------------------------------------------------------
RONG_CM = 15.5
CM_KY_TU = 0.19       # bề rộng trung bình một ký tự
CAO_DONG_CM = 0.45    # chiều cao một dòng chữ


def _ten_chuoi(hinh):
    cfg = hinh["bieu_do"]
    return [hinh["cot"][s["cot"]] for s in cfg.get("chuoi", [])]


def _so_dong_chu_giai(hinh):
    cfg = hinh["bieu_do"]
    if cfg.get("an_chu_giai") or cfg["loai"] == "doughnut":
        return 0
    rong_kha_dung = RONG_CM - 0.6
    dong, x = 1, 0.0
    for ten in _ten_chuoi(hinh):
        w = len(ten) * CM_KY_TU + 0.9
        if x > 0 and x + w > rong_kha_dung:
            dong += 1
            x = 0.0
        x += w
    return dong


def _dong_nhan(text, rong_cm):
    """Số dòng khi một nhãn danh mục được ngắt trong bề rộng rong_cm."""
    import math
    tu = str(text).split()
    dong, x = 1, 0.0
    for t in tu:
        w = (len(t) + 1) * CM_KY_TU
        if x > 0 and x + w > rong_cm:
            dong += 1
            x = 0.0
        x += w
    return max(1, dong) if tu else 1


def bo_cuc(hinh, cao_cm=None):
    """Trả về (chiều cao đề xuất, layout vùng vẽ, layout chú giải) theo tỷ lệ 0..1."""
    cfg = hinh["bieu_do"]
    loai = cfg["loai"]
    dm = [d[0] for d in hinh["dong"]]
    co_y2 = any(s.get("truc_phu") for s in cfg.get("chuoi", []))
    n_cg = _so_dong_chu_giai(hinh)
    cao_cg = n_cg * 0.5 + 0.15 if n_cg else 0.0
    tren = cao_cg + (0.35 if n_cg else 0.3)

    if loai == "doughnut":
        cao = cao_cm or 7.0
        return cao, None, None

    if loai == "bar":
        rong_nhan = min(5.6, max(len(str(x)) for x in dm) * CM_KY_TU + 0.5)
        so_dong = [_dong_nhan(x, rong_nhan - 0.3) for x in dm]
        so_hang = len(cfg.get("chuoi", [])) if not cfg.get("xep_chong") else 1
        cao_hang = [max(0.55 + 0.32 * (so_hang - 1), k * CAO_DONG_CM + 0.25) for k in so_dong]
        duoi = 0.55 + (0.55 if cfg.get("truc_y") else 0)
        cao = cao_cm or round(tren + sum(cao_hang) + duoi + 0.2, 1)
        trai = rong_nhan + 0.25
        phai = 1.1
        plot = {"x": trai / RONG_CM, "y": tren / cao,
                "width": (RONG_CM - trai - phai) / RONG_CM, "height": (cao - tren - duoi) / cao}
    else:
        cao = cao_cm or 8.5 + 0.5 * max(0, n_cg - 1)
        trai = 1.55 if cfg.get("truc_y") else 1.0
        phai = 1.75 if co_y2 else 0.45
        rong_ve = RONG_CM - trai - phai
        o = rong_ve / max(1, len(dm))
        dong_dm = max(_dong_nhan(x, o - 0.15) for x in dm)
        duoi = dong_dm * CAO_DONG_CM + 0.35 + (0.6 if cfg.get("truc_x") else 0)
        plot = {"x": trai / RONG_CM, "y": tren / cao,
                "width": rong_ve / RONG_CM, "height": (cao - tren - duoi) / cao}
    cg = None
    if n_cg:
        cg = {"x": 0.02, "y": 0.1 / cao, "width": 0.96, "height": cao_cg / cao}
    return cao, plot, cg


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
            "data_labels": {"percentage": True, "value": True, "separator": "\n", "font": _font(11, True),
                            "custom": [{"font": _font(11, True, mau_chu_tren(m))} for m in cfg["mau_diem"]]},
        })
        ch.set_hole_size(cfg.get("lo", 55))
        ch.set_legend({"position": "right", "font": _font(10),
                       "layout": {"x": 0.62, "y": 0.3, "width": 0.36, "height": 0.4}})
        ch.set_plotarea({"border": {"none": True},
                         "layout": {"x": 0.08, "y": 0.06, "width": 0.5, "height": 0.88}})
        _chung(ch, tieu_de, co_plot=False)
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
                dl["font"] = _font(9, True, mau_chu_tren(s["mau"]))
            if kieu in ("column", "bar") and dl.get("position") == "center" and s.get("mau_diem"):
                dl["custom"] = [{"font": _font(9, True, mau_chu_tren(m))} for m in s["mau_diem"]]
            if s.get("an_nhan_0"):
                cu = dl.get("custom") or [None] * len(hinh["dong"])
                dl["custom"] = [{"delete": True} if (d[s["cot"]] or 0) == 0 else cu[i]
                                for i, d in enumerate(hinh["dong"])]
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
    if cfg.get("dd_y"):
        truc_gia_tri["num_format"] = cfg["dd_y"]
    if cfg.get("max_y"):
        truc_gia_tri["max"] = cfg["max_y"]
    if cfg.get("buoc_y"):
        truc_gia_tri["major_unit"] = cfg["buoc_y"]
    if cfg.get("truc_y"):
        truc_gia_tri["name"] = cfg["truc_y"]
    truc_danh_muc = {"num_font": _font(10), "line": {"color": MAU_TRUC, "width": 0.75},
                     "major_tick_mark": "none", "minor_tick_mark": "none",
                     "interval_unit": 1}  # hiện đủ mọi nhãn danh mục, không để phần mềm tự bỏ bớt
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

    _, plot, cg = bo_cuc(hinh, hinh.get("cao"))
    if cfg.get("an_chu_giai"):
        ch.set_legend({"none": True})
    else:
        lg = {"position": "top", "font": _font(10)}
        if cg:
            lg["layout"] = cg
        ch.set_legend(lg)
    ch.set_plotarea({"border": {"none": True}, "layout": plot})
    _chung(ch, tieu_de, co_plot=False)
    return ch


def _chung(ch, tieu_de, co_plot=True):
    # Tiêu đề hình đặt trong Word (dòng "Hình 2.x"), không đặt trong biểu đồ để không chiếm chỗ vùng vẽ.
    ch.set_title({"none": True})
    ch.set_chartarea({"border": {"none": True}, "fill": {"color": "#FFFFFF"}})
    if co_plot:
        ch.set_plotarea({"border": {"none": True}})
