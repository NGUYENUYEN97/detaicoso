# -*- coding: utf-8 -*-
"""Sửa văn bản Word ở chế độ theo dõi thay đổi (Track Changes).

Mỗi chỉnh sửa là cặp (chuỗi cũ, chuỗi mới). Chuỗi cũ phải xuất hiện đúng một lần
trong toàn văn bản (kể cả trong bảng). Phần bị thay được đánh dấu xóa, phần mới
được đánh dấu chèn, người đọc có thể chấp nhận hoặc từ chối từng chỗ trong Word.
"""
import copy
import datetime
import os
import shutil
import subprocess
import tempfile
import zipfile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"


def q(tag):
    return f"{{{W}}}{tag}"


class TrackEditor:
    def __init__(self, root, author, date=None):
        self.root = root
        self.author = author
        self.date = date or datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        ids = [int(x) for x in root.xpath("//@w:id", namespaces={"w": W}) if str(x).lstrip("-").isdigit()]
        self.next_id = max(ids + [0]) + 1000

    def _id(self):
        self.next_id += 1
        return str(self.next_id)

    @staticmethod
    def _runs(p):
        return [r for r in p.findall(q("r"))]

    @staticmethod
    def _text(r):
        return "".join(t.text or "" for t in r.findall(q("t")))

    def _para_text(self, p):
        return "".join(self._text(r) for r in self._runs(p))

    def _split(self, r, pos):
        """Tách run r tại vị trí ký tự pos; trả về (run trái, run phải)."""
        text = self._text(r)
        left, right = text[:pos], text[pos:]
        r2 = copy.deepcopy(r)
        for rr, tx in ((r, left), (r2, right)):
            ts = rr.findall(q("t"))
            for t in ts[1:]:
                rr.remove(t)
            if ts:
                ts[0].text = tx
                ts[0].set(f"{{{XML_NS}}}space", "preserve")
        r.addnext(r2)
        return r, r2

    def replace(self, old, new, tat_ca=False):
        hits = [p for p in self.root.iter(q("p")) if old in self._para_text(p)]
        if not hits or (len(hits) != 1 and not tat_ca):
            raise ValueError(f"Tìm thấy {len(hits)} đoạn chứa: {old[:90]}")
        so_lan = 0
        while hits:
            self._replace_in(hits[0], old, new)
            so_lan += 1
            hits = [p for p in self.root.iter(q("p")) if old in self._para_text(p)] if tat_ca else []
        return so_lan

    def _replace_in(self, p, old, new):
        full = self._para_text(p)
        start = full.index(old)
        end = start + len(old)
        # tách run tại start và end
        for cut in (end, start):
            pos = 0
            for r in self._runs(p):
                n = len(self._text(r))
                if pos < cut < pos + n:
                    self._split(r, cut - pos)
                    break
                pos += n
        pos = 0
        muc_tieu = []
        for r in self._runs(p):
            n = len(self._text(r))
            if start <= pos and pos + n <= end and n > 0:
                muc_tieu.append(r)
            pos += n
        if not muc_tieu:
            raise ValueError(f"Không cô lập được đoạn chữ: {old[:60]}")
        rpr = muc_tieu[0].find(q("rPr"))
        d = etree.Element(q("del"))
        d.set(q("id"), self._id())
        d.set(q("author"), self.author)
        d.set(q("date"), self.date)
        muc_tieu[0].addprevious(d)
        for r in muc_tieu:
            for t in r.findall(q("t")):
                t.tag = q("delText")
            d.append(r)
        if new:
            ins = etree.Element(q("ins"))
            ins.set(q("id"), self._id())
            ins.set(q("author"), self.author)
            ins.set(q("date"), self.date)
            nr = etree.SubElement(ins, q("r"))
            if rpr is not None:
                nr.append(copy.deepcopy(rpr))
            for k, dong in enumerate(new.split("\n")):
                if k:
                    etree.SubElement(nr, q("br"))
                t = etree.SubElement(nr, q("t"))
                t.text = dong
                t.set(f"{{{XML_NS}}}space", "preserve")
            d.addnext(ins)


    def _doan_duy_nhat(self, neo):
        hits = [p for p in self.root.iter(q("p")) if neo in self._para_text(p)]
        if len(hits) != 1:
            raise ValueError(f"Tìm thấy {len(hits)} đoạn chứa: {neo[:90]}")
        return hits[0]

    def them_doan_sau(self, neo, text, mau=None):
        """Chèn một đoạn mới ngay sau đoạn neo, đánh dấu là nội dung chèn.

        neo: chuỗi nằm trong đúng một đoạn, hoặc chính phần tử đoạn. mau: chuỗi xác định đoạn lấy định dạng
        (mặc định lấy định dạng của đoạn neo). Trả về phần tử đoạn mới để chèn nối tiếp."""
        p = neo if not isinstance(neo, str) else self._doan_duy_nhat(neo)
        mau_p = self._doan_duy_nhat(mau) if mau else p
        moi = etree.Element(q("p"))
        ppr = mau_p.find(q("pPr"))
        if ppr is not None:
            ppr = copy.deepcopy(ppr)
            moi.append(ppr)
        else:
            ppr = etree.SubElement(moi, q("pPr"))
        rpr_p = ppr.find(q("rPr"))
        if rpr_p is None:
            rpr_p = etree.SubElement(ppr, q("rPr"))
        dau = etree.Element(q("ins"))
        rpr_p.insert(0, dau)  # w:ins phải đứng đầu rPr của dấu đoạn
        dau.set(q("id"), self._id())
        dau.set(q("author"), self.author)
        dau.set(q("date"), self.date)
        ins = etree.SubElement(moi, q("ins"))
        ins.set(q("id"), self._id())
        ins.set(q("author"), self.author)
        ins.set(q("date"), self.date)
        nr = etree.SubElement(ins, q("r"))
        runs = self._runs(mau_p) or [r for r in mau_p.iter(q("r"))]
        if runs and runs[0].find(q("rPr")) is not None:
            nr.append(copy.deepcopy(runs[0].find(q("rPr"))))
        t = etree.SubElement(nr, q("t"))
        t.text = text
        t.set(f"{{{XML_NS}}}space", "preserve")
        p.addnext(moi)
        return moi


def ap_dung(vao, ra, sua, author, merge_runs=None):
    tmp = tempfile.mkdtemp()
    src = os.path.join(tmp, "src.docx")
    shutil.copy(vao, src)
    if merge_runs:
        subprocess.run(["python3", merge_runs, src, "-o", src], check=True, capture_output=True)
    with zipfile.ZipFile(src) as z:
        xml = z.read("word/document.xml")
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    root = etree.fromstring(xml)
    ed = TrackEditor(root, author)
    for muc in sua:
        cu, moi = muc[0], muc[1]
        if len(muc) > 2 and muc[2] == "doan_moi_sau":
            ed.them_doan_sau(cu, moi)
            continue
        if len(muc) > 2 and muc[2] == "cac_doan_sau":
            # moi là danh sách (chuỗi xác định đoạn mẫu định dạng hoặc None, nội dung)
            el = cu
            for mau, text in moi:
                el = ed.them_doan_sau(el, text, mau=mau)
            continue
        ed.replace(cu, moi, tat_ca=len(muc) > 2 and muc[2] == "tat_ca")
    moi_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(ra, "w", zipfile.ZIP_DEFLATED) as zout:
        for info, data in items:
            zout.writestr(info, moi_xml if info.filename == "word/document.xml" else data)
    shutil.rmtree(tmp)
    return len(sua)
