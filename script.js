const produk = [
  { kode: 'P001', nama: 'Speaker Mini', stokSistem: 10 },
  { kode: 'P002', nama: 'Headset Bluetooth', stokSistem: 7 },
  { kode: 'P003', nama: 'Kabel USB-C', stokSistem: 15 },
];

const tbody = document.getElementById("produk-table");

produk.forEach((item, i) => {
  const tr = document.createElement("tr");

  tr.innerHTML = `
    <td>${item.kode}</td>
    <td>${item.nama}</td>
    <td>${item.stokSistem}</td>
    <td><input type="number" id="fisik-${i}" value="0" min="0"></td>
    <td id="selisih-${i}">0</td>
  `;

  tbody.appendChild(tr);

  const input = tr.querySelector(`#fisik-${i}`);
  input.addEventListener("input", () => {
    const fisik = parseInt(input.value) || 0;
    const selisih = fisik - item.stokSistem;
    document.getElementById(`selisih-${i}`).textContent = selisih;
  });
});

function exportToExcel() {
  const dataExport = produk.map((item, i) => {
    const fisikInput = document.getElementById(`fisik-${i}`);
    const stokFisik = parseInt(fisikInput.value) || 0;
    const selisih = stokFisik - item.stokSistem;

    return {
      "Kode Produk": item.kode,
      "Nama Produk": item.nama,
      "Stok Sistem": item.stokSistem,
      "Stok Fisik": stokFisik,
      "Selisih": selisih,
    };
  });

  const worksheet = XLSX.utils.json_to_sheet(dataExport);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Laporan Opname");

  XLSX.writeFile(workbook, "Laporan_Stock_Opname.xlsx");
}
