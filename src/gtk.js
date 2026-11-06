/* ============ ANGGOTA GTK (Guru dan Tenaga Kependidikan) ============
   Desain dan alur sama dengan Anggota Siswa. Koleksi Firestore: anggota_gtk (ID dokumen = No. Anggota).
   No. Anggota adalah isi barcode kartu yang di-scan pada menu Peminjaman dan Pengembalian. */
import * as XLSX from 'xlsx';

export function setupGtk(c) {
  const { S, Q, SEL, views, ACT, INP, CHG, SUB, PER, $, esc, txt, norm, ic, natSort, emptyHtml, pagerHtml, mHead, openModal, closeModal,
    confirmBox, toast, withProg, quick, commit, actOp, activeOf, fmtNum, validKode, doPrint, LOGO, openImport, needXlsx, xl, sheetFrom,
    blankText, saveWb, stamp } = c;
  const JENIS = ['Guru', 'Tendik'];

  function gtkFiltered() {
    const q = norm(Q.g.q);
    return S.gtk.filter(g => {
      if (Q.g.jenis && g.jenisPtk !== Q.g.jenis) return false;
      if (q) return norm(g.nama).includes(q) || norm(g.noAnggota).includes(q) || norm(g.nip).includes(q);
      return true;
    }).sort((a, b) => natSort(a.jenisPtk, b.jenisPtk) || natSort(a.nama, b.nama));
  }
  function gtkPage() {
    const all = gtkFiltered(), pages = Math.max(1, Math.ceil(all.length / PER));
    if (Q.g.page > pages) Q.g.page = pages;
    return { all, rows: all.slice((Q.g.page - 1) * PER, Q.g.page * PER) };
  }
  function updSelG() {
    const n = [...SEL.gtk].filter(id => S.gtk.some(g => g.id === id)).length;
    const del = $('#g-del'), card = $('#g-card'), ha = $('[data-ch=selAllGtk]');
    if (del) { del.style.display = n ? '' : 'none'; del.innerHTML = `${ic('trash', 16)}Hapus terpilih (${n})`; }
    if (card) card.innerHTML = `${ic('card', 18)}Cetak Kartu${n ? ' (' + n + ')' : ''}`;
    if (ha) { const ids = gtkPage().rows.map(g => g.id); ha.checked = ids.length > 0 && ids.every(i => SEL.gtk.has(i)); }
  }

  views.gtk = {
    mount(el) {
      el.innerHTML = `
<div class="vh"><div><h2>Anggota GTK</h2><p class="sub">Data guru dan tenaga kependidikan sebagai anggota perpustakaan. No. Anggota dipakai sebagai barcode kartu.</p></div></div>
<div class="tools">
  <div class="srch">${ic('search', 18)}<input class="in" data-in="gq" placeholder="Cari nama, No. Anggota, atau NIP\u2026" value="${esc(Q.g.q)}" aria-label="Cari anggota GTK"></div>
  <select class="in" data-ch="gjenis" aria-label="Jenis PTK"><option value="">Semua jenis PTK</option>${JENIS.map(j => `<option>${j}</option>`).join('')}</select>
</div>
<div class="tools">
  <button class="btn pri" data-act="gtkAdd">${ic('plus', 18)}Tambah Anggota GTK</button>
  <button class="btn" data-act="importGtk">${ic('upload', 18)}Impor xlsx</button>
  <button class="btn" data-act="exportGtk">${ic('download', 18)}Ekspor xlsx</button>
  <button class="btn" data-act="tplGtk">${ic('file', 18)}Template</button>
  <button class="btn warn" data-act="gtkKartuOpen" id="g-card">${ic('card', 18)}Cetak Kartu</button>
  <button class="btn dng sm" data-act="gtkDelSel" id="g-del" style="display:none">${ic('trash', 16)}Hapus terpilih</button>
</div>
<div id="g-zone"></div>`;
    },
    update() {
      const sel = $('[data-ch=gjenis]'); if (sel) sel.value = Q.g.jenis;
      const { all, rows } = gtkPage(), z = $('#g-zone'); if (!z) return;
      if (!all.length) { z.innerHTML = `<div class="tw">${emptyHtml('users', S.gtk.length ? 'Tidak ada anggota GTK yang cocok.' : 'Belum ada anggota GTK. Klik "Tambah Anggota GTK" atau impor dari file xlsx.')}</div>`; updSelG(); return; }
      z.innerHTML = `<div class="tw"><table class="tbl"><thead><tr><th><input type="checkbox" data-ch="selAllGtk" aria-label="Pilih semua di halaman ini"></th><th>No. Anggota</th><th>NIP</th><th>Nama</th><th>Jenis PTK</th><th>L/P</th><th>Dipinjam</th><th></th></tr></thead><tbody>${rows.map(g => {
        const n = activeOf(g.noAnggota).length;
        return `<tr><td data-l=""><input type="checkbox" data-ch="selGtk" data-id="${esc(g.id)}" ${SEL.gtk.has(g.id) ? 'checked' : ''} aria-label="Pilih ${esc(g.nama)}"></td>
<td data-l="No. Anggota"><b>${esc(g.noAnggota)}</b></td><td data-l="NIP">${esc(g.nip || '-')}</td><td data-l="Nama"><div>${esc(g.nama)}${g.aktif === false ? ' <span class="bdg mut">Nonaktif</span>' : ''}</div></td>
<td data-l="Jenis PTK">${esc(g.jenisPtk)}</td><td data-l="L/P">${esc(g.jenisKelamin || '-')}</td><td data-l="Dipinjam">${n ? `<span class="bdg warn">${n} buku</span>` : '<span class="bdg mut">0</span>'}</td>
<td class="act" data-l=""><button class="ibtn" data-act="gtkKartuOne" data-id="${esc(g.id)}" title="Cetak kartu" aria-label="Cetak kartu">${ic('card', 18)}</button><button class="ibtn" data-act="gtkEdit" data-id="${esc(g.id)}" title="Ubah" aria-label="Ubah">${ic('edit', 18)}</button><button class="ibtn d" data-act="gtkDel" data-id="${esc(g.id)}" title="Hapus" aria-label="Hapus">${ic('trash', 18)}</button></td></tr>`;
      }).join('')}</tbody></table></div>${pagerHtml(all.length, Q.g.page, PER)}`;
      updSelG();
    }
  };
  INP.gq = t => { Q.g.q = t.value; Q.g.page = 1; views.gtk.update(); };
  CHG.gjenis = t => { Q.g.jenis = t.value; Q.g.page = 1; views.gtk.update(); };
  CHG.selGtk = t => { t.checked ? SEL.gtk.add(t.dataset.id) : SEL.gtk.delete(t.dataset.id); updSelG(); };
  CHG.selAllGtk = t => { gtkPage().rows.forEach(g => t.checked ? SEL.gtk.add(g.id) : SEL.gtk.delete(g.id)); views.gtk.update(); };

  /* ---- form tambah / ubah ---- */
  function gtkForm(g) {
    const edit = !!g; g = g || { jenisPtk: 'Guru', jenisKelamin: 'L', aktif: true };
    openModal(`${mHead(edit ? 'Ubah Anggota GTK' : 'Tambah Anggota GTK')}
<form data-sub="gtkSave" data-edit="${edit ? 1 : 0}" autocomplete="off" style="display:contents">
<div class="mb"><div class="grid2">
 <div class="fld"><label>No. Anggota (isi barcode kartu)*</label><input class="in" name="no" value="${esc(g.noAnggota || '')}" ${edit ? 'readonly' : ''} required maxlength="30" placeholder="contoh: G001"></div>
 <div class="fld"><label>NIP</label><input class="in" name="nip" value="${esc(g.nip || '')}" maxlength="20" inputmode="numeric"><small>Boleh dikosongkan bila belum punya NIP. Disimpan sebagai teks.</small></div>
 <div class="fld full"><label>Nama lengkap*</label><input class="in" name="nama" value="${esc(g.nama || '')}" required maxlength="120"></div>
 <div class="fld"><label>Jenis PTK*</label><select class="in" name="jenis" style="width:100%">${JENIS.map(j => `<option value="${j}" ${g.jenisPtk === j ? 'selected' : ''}>${j}</option>`).join('')}</select></div>
 <div class="fld"><label>Jenis kelamin*</label><select class="in" name="jk" style="width:100%"><option value="L" ${g.jenisKelamin === 'L' ? 'selected' : ''}>Laki-laki</option><option value="P" ${g.jenisKelamin === 'P' ? 'selected' : ''}>Perempuan</option></select></div>
 <label class="chk full"><input type="checkbox" name="aktif" ${g.aktif !== false ? 'checked' : ''}> Anggota aktif (boleh meminjam)</label>
</div></div>
<div class="mf"><button type="button" class="btn" data-act="closeModal">Batal</button><button class="btn pri" type="submit">${ic('check', 18)}Simpan</button></div></form>`);
  }
  ACT.gtkAdd = () => gtkForm();
  ACT.gtkEdit = t => { const g = S.gtk.find(x => x.id === t.dataset.id); if (g) gtkForm(g); };
  SUB.gtkSave = async f => {
    const fd = new FormData(f), edit = f.dataset.edit === '1';
    const no = txt(fd.get('no')), nip = txt(fd.get('nip')), nama = txt(fd.get('nama')), jenis = JENIS.includes(fd.get('jenis')) ? fd.get('jenis') : 'Guru', jk = fd.get('jk') === 'P' ? 'P' : 'L';
    let err = validKode(no).replace('Kode', 'No. Anggota');
    if (!err && !nama) err = 'Nama wajib diisi.';
    if (!err && nip && !/^\d+$/.test(nip)) err = 'NIP hanya boleh berisi angka.';
    if (!err && !edit && S.aMap.has(norm(no))) err = 'No. Anggota sudah dipakai (oleh siswa atau GTK lain).';
    if (!err && nip) { const o = S.gtk.find(x => norm(x.nip) === norm(nip) && norm(x.noAnggota) !== norm(no)); if (o) err = 'NIP sudah dipakai ' + o.nama + '.'; }
    if (err) { toast(err, 'err'); return; }
    const d = { noAnggota: no, nip, nama, jenisPtk: jenis, jenisKelamin: jk, aktif: !!fd.get('aktif') };
    const ops = [];
    if (edit) { const old = S.gtk.find(x => norm(x.noAnggota) === norm(no)); if (!old) { toast('Data anggota GTK tidak ditemukan.', 'err'); return; } ops.push({ t: 'upd', c: 'anggota_gtk', id: old.id, d }); }
    else { d.dibuatPada = Date.now(); ops.push({ t: 'set', c: 'anggota_gtk', id: no, d }); ops.push(actOp('anggota', nama, jenis)); }
    const btn = $('button[type=submit]', f); btn.disabled = true;
    try { await withProg(edit ? 'Memperbarui anggota GTK' : 'Menyimpan anggota GTK', nama, () => commit(ops)); closeModal(); toast(edit ? 'Data anggota GTK diperbarui.' : 'Anggota GTK ' + nama + ' ditambahkan.', 'ok'); }
    catch (e) { btn.disabled = false; toast('Gagal menyimpan: ' + (e.message || e), 'err', 5000); }
  };
  ACT.gtkDel = async t => {
    const g = S.gtk.find(x => x.id === t.dataset.id); if (!g) return;
    if (activeOf(g.noAnggota).length) { toast('Anggota masih memiliki pinjaman aktif.', 'err'); return; }
    if (!await confirmBox(`Hapus anggota GTK ${g.nama} (${g.noAnggota})? Riwayat peminjaman tetap tersimpan.`, 'Hapus', true)) return;
    try { await quick(() => commit([{ t: 'del', c: 'anggota_gtk', id: g.id }])); SEL.gtk.delete(g.id); toast('Anggota GTK dihapus.', 'ok'); } catch (e) { toast('Gagal menghapus: ' + (e.message || e), 'err'); }
  };
  ACT.gtkDelSel = async () => {
    const items = [...SEL.gtk].map(id => S.gtk.find(g => g.id === id)).filter(Boolean);
    const del = items.filter(g => !activeOf(g.noAnggota).length), skip = items.length - del.length;
    if (!del.length) { toast('Semua anggota terpilih masih memiliki pinjaman aktif.', 'err'); return; }
    if (!await confirmBox(`Hapus ${del.length} anggota GTK terpilih?${skip ? ' ' + skip + ' anggota dengan pinjaman aktif dilewati.' : ''}`, 'Hapus', true)) return;
    try { await withProg('Menghapus anggota GTK', del.length + ' anggota', p => commit(del.map(g => ({ t: 'del', c: 'anggota_gtk', id: g.id })), (d, t) => p.set(d / t * 100, fmtNum(d) + ' dari ' + fmtNum(t)))); del.forEach(g => SEL.gtk.delete(g.id)); toast(del.length + ' anggota GTK dihapus.', 'ok'); } catch (e) { toast('Gagal menghapus: ' + (e.message || e), 'err'); }
  };

  /* ---- cetak kartu (ukuran dan tata letak sama dengan kartu siswa) ---- */
  let KG = { ids: null };
  function kartuModalG(ids) {
    KG.ids = ids || null;
    const flt = gtkFiltered().length, sel = [...SEL.gtk].filter(id => S.gtk.some(g => g.id === id)).length, opts = [];
    if (ids) opts.push(['ids', 'Anggota ini saja', ids.length]);
    else { if (sel) opts.push(['sel', 'Anggota yang dicentang', sel]); opts.push(['flt', 'Hasil pencarian/filter saat ini', flt]); }
    opts.push(['all', 'Semua anggota GTK', S.gtk.length]);
    openModal(`${mHead('Cetak Kartu Perpustakaan GTK')}
<form data-sub="gtkKartuPrint" style="display:contents"><div class="mb">
 <div class="fld"><label>Anggota yang dicetak</label>${opts.map((o, i) => `<label class="chk"><input type="radio" name="src" value="${o[0]}" ${i === 0 ? 'checked' : ''}> ${o[1]} (${fmtNum(o[2])})</label>`).join('')}</div>
 <p class="hint" style="margin-top:12px">Ukuran kartu 85,6 x 54 mm, 10 kartu per lembar A4. Barcode berisi No. Anggota. Pada kotak cetak browser aktifkan "Grafik latar belakang" dan skala 100%.</p>
</div>
<div class="mf"><button type="button" class="btn" data-act="closeModal">Batal</button><button class="btn warn" type="submit">${ic('print', 18)}Cetak Kartu</button></div></form>`);
  }
  ACT.gtkKartuOpen = () => kartuModalG(null);
  ACT.gtkKartuOne = t => kartuModalG([t.dataset.id]);
  SUB.gtkKartuPrint = async f => {
    const src = new FormData(f).get('src'); let list;
    if (src === 'ids') list = KG.ids.map(id => S.gtk.find(g => g.id === id)).filter(Boolean);
    else if (src === 'sel') list = [...SEL.gtk].map(id => S.gtk.find(g => g.id === id)).filter(Boolean);
    else if (src === 'flt') list = gtkFiltered();
    else list = S.gtk.slice().sort((a, b) => natSort(a.jenisPtk, b.jenisPtk) || natSort(a.nama, b.nama));
    if (!list.length) { toast('Tidak ada anggota untuk dicetak.', 'err'); return; }
    if (list.length > 1000) { toast('Maksimal 1.000 kartu sekali cetak. Cetak per jenis PTK.', 'err'); return; }
    const logo = S.set.logoUrl || LOGO;
    const html = `<div class="kt-grid">${list.map(g => `<div class="kt"><div class="kh"><img src="${esc(logo)}" alt=""><div><b>${esc(S.set.namaSekolah)}</b><span>KARTU PERPUSTAKAAN GTK</span></div></div><div class="kn"><b>${esc(g.nama)}</b><span>${esc(g.jenisPtk)} \u2022 No. ${esc(g.noAnggota)}</span></div><div class="kb"><svg data-bc="${esc(g.noAnggota)}" data-h="40" data-fs="12"></svg></div><div class="ks"></div></div>`).join('')}</div>`;
    closeModal();
    await doPrint(html, 'Kartu perpustakaan GTK ' + fmtNum(list.length) + ' kartu');
  };

  /* ---- impor / ekspor / template xlsx ---- */
  ACT.importGtk = () => { if (needXlsx()) openImport('gtk'); };
  ACT.exportGtk = () => xl('Mengekspor data anggota GTK', () => {
    const rows = S.gtk.slice().sort((a, b) => natSort(a.jenisPtk, b.jenisPtk) || natSort(a.nama, b.nama));
    const aoa = [['No. Anggota', 'NIP', 'Nama', 'Jenis PTK', 'Jenis Kelamin', 'Status']].concat(rows.map(g => [txt(g.noAnggota), txt(g.nip), txt(g.nama), txt(g.jenisPtk), txt(g.jenisKelamin), g.aktif === false ? 'Nonaktif' : 'Aktif']));
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, sheetFrom(aoa, [0, 1], [16, 22, 36, 12, 14, 10]), 'Anggota GTK');
    saveWb(wb, 'anggota_gtk_' + stamp() + '.xlsx');
  });
  ACT.tplGtk = () => xl('Menyiapkan template anggota GTK', () => {
    const aoa = [['No. Anggota', 'NIP', 'Nama', 'Jenis PTK', 'Jenis Kelamin'], ['G001', '198001012005012001', 'Contoh Guru', 'Guru', 'P'], ['T001', '', 'Contoh Tendik', 'Tendik', 'L']];
    const ws = sheetFrom(aoa, [0, 1], [16, 22, 36, 12, 14]); blankText(ws, 3, 300, [0, 1]);
    const info = XLSX.utils.aoa_to_sheet([['PETUNJUK PENGISIAN'], ['1. Isi data mulai baris ke-2 pada sheet "Anggota GTK". Jangan mengubah judul kolom.'], ['2. No. Anggota wajib unik, tanpa spasi dan tanda /. Nomor ini menjadi isi barcode kartu yang di-scan.'], ['3. No. Anggota dan NIP berformat Teks. Jangan diubah menjadi Angka agar angka 0 di depan tidak hilang.'], ['4. Jenis PTK diisi Guru atau Tendik. Jenis Kelamin diisi L (laki-laki) atau P (perempuan).'], ['5. NIP boleh dikosongkan.']]);
    info['!cols'] = [{ wch: 110 }];
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Anggota GTK'); XLSX.utils.book_append_sheet(wb, info, 'Petunjuk');
    saveWb(wb, 'template_anggota_gtk.xlsx');
  });
}
