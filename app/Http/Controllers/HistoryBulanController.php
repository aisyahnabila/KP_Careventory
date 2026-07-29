<?php

namespace App\Http\Controllers;

use App\Models\Permintaan;
use App\Models\UnitKerja;
use Illuminate\Http\Request;
use Illuminate\Support\Carbon;
use PhpOffice\PhpWord\TemplateProcessor;

class HistoryBulanController extends Controller
{
    public function index(Request $request)
    {
        $unit_kerja = $request->input('unit_kerja');
        $bulan = $request->input('bulan');
        $tahun = $request->input('tahun');

        $query = Permintaan::query();

        if ($unit_kerja) {
            $query->where('id_unitkerja', $unit_kerja);
        }

        if ($bulan) {
            $query->whereMonth('tanggal_permintaan', $bulan);
        }

        if ($tahun) {
            $query->whereYear('tanggal_permintaan', $tahun);
        }

        // Get all permintaan with related models
        $permintaan = $query->with('detailPermintaan.barang.kategori', 'unitKerja')->get();

        // Group and transform data
        $allDetails = $permintaan->flatMap(function ($p) {
            return $p->detailPermintaan;
        });

        $groupedDetails = $allDetails->groupBy(function ($detail) {
            $unitKerjaName = $detail->permintaan->unitKerja->nama_unit_kerja ?? 'N/A';
            $bulan = \Carbon\Carbon::parse($detail->permintaan->tanggal_permintaan)->translatedFormat('F Y');
            $namaBarang = optional($detail->barang)->nama_barang ?? 'N/A';
            $specBarang = optional($detail->barang)->spesifikasi_nama_barang ?? 'N/A';
            return $unitKerjaName . '-' . $bulan . '-' . $namaBarang . '-' . $specBarang;
        });

        $reportData = $groupedDetails->map(function ($group) {
            $firstDetail = $group->first();
            $totalPermintaan = $group->sum('jumlah_permintaan');
            $barang = optional($firstDetail->barang);
            $permintaan = optional($firstDetail->permintaan);
            $unitKerja = optional($permintaan->unitKerja);
            $kategori = optional($barang->kategori);

            return [
                'bulan' => $permintaan->tanggal_permintaan ? \Carbon\Carbon::parse($permintaan->tanggal_permintaan)->translatedFormat('F Y') : 'N/A',
                'tanggal_permintaan_raw' => $permintaan->tanggal_permintaan, // Add this for sorting
                'unit_kerja' => $unitKerja->nama_unit_kerja ?? 'N/A',
                'kode_barang' => $kategori->kode_barang ?? 'N/A',
                'nama_barang' => $barang->nama_barang ?? 'N/A',
                'spesifikasi_nama_barang' => $barang->spesifikasi_nama_barang ?? 'N/A',
                'total_permintaan' => $totalPermintaan,
                'jumlah' => $barang->jumlah ?? 0,
                'satuan' => $barang->satuan ?? 'N/A',
                'keperluan' => $permintaan->keperluan ?? 'N/A',
            ];
        });

        // Sort the report data by 'tanggal_permintaan_raw' in descending order (newest first)
        $reportData = $reportData->sortByDesc('tanggal_permintaan_raw')->values();

        return view('laporan.bulan', [
            'permintaan' => $reportData,
            'unitKerjaOptions' => UnitKerja::all(),
        ]);
    }

    public function exportToWord(Request $request)
    {
        $htmlFlags = ENT_XML1 | ENT_QUOTES;
        // path ke template
        $templatePath = public_path('templates/template_surat_permintaan_barang.docx');

        // buat instance TemplateProcessor
        $templateProcessor = new TemplateProcessor($templatePath);

        // mengambil data dari filter
        $unit_kerja = $request->input('unit_kerja');
        $bulan = $request->input('bulan');
        $tahun = $request->input('tahun');

        // simpan data yang sudah difilter dari database
        $permintaan = Permintaan::when($unit_kerja, function ($query, $unit_kerja) {
            return $query->where('id_unitkerja', $unit_kerja);
        })
            ->when($bulan, function ($query, $bulan) {
                return $query->whereMonth('tanggal_permintaan', $bulan);
            })
            ->when($tahun, function ($query, $tahun) {
                return $query->whereYear('tanggal_permintaan', $tahun);
            })
            ->with('detailPermintaan.barang.kategori', 'unitKerja')
            ->get();

        // Group and aggregate data
        $details = $permintaan->flatMap->detailPermintaan;
        $groupedDetails = $details->groupBy(function ($detail) {
            return $detail->permintaan->unitKerja->nama_unit_kerja . '-' . \Carbon\Carbon::parse($detail->permintaan->tanggal_permintaan)->format('F Y') . '-' . $detail->barang->nama_barang . '-' . $detail->barang->spesifikasi_nama_barang;
        })->map(function ($group) {
            $first = $group->first();
            $first->jumlah_permintaan = $group->sum(function ($detail) {
                return (int) $detail->jumlah_permintaan;
            });
            return $first;
        });

        // mengisi placeholder dengan data yang sudah difilter
        $templateProcessor->setValue('unit_kerja', htmlspecialchars(optional($permintaan->first())->unitKerja->nama_unit_kerja ?? 'Semua Unit Kerja', $htmlFlags, 'UTF-8'));
        $templateProcessor->setValue('bulan', htmlspecialchars($bulan, $htmlFlags, 'UTF-8'));
        $templateProcessor->setValue('tahun', htmlspecialchars($tahun, $htmlFlags, 'UTF-8'));

        // menambahkan tanggal cetak
        $tanggalCetak = Carbon::now()->format('d-m-Y');
        $templateProcessor->setValue('tanggal_cetak', $tanggalCetak);

        // loop through grouped details and fill in the table
        $templateProcessor->cloneRow('kode_barang', $groupedDetails->count());
        foreach ($groupedDetails->values() as $index => $detail) {
            $index = $index + 1; // Adjusting index to start from 1

            $stok_awal = $detail->barang->jumlah + $detail->jumlah_permintaan;
            $sisa_persediaan = $stok_awal - $detail->jumlah_permintaan;
            $usulan_pengajuan_persetujuan = $detail->jumlah_permintaan;

            $templateProcessor->setValue("no#{$index}", htmlspecialchars($index, $htmlFlags, 'UTF-8'));
            $templateProcessor->setValue("unit_kerja#{$index}", htmlspecialchars($detail->permintaan->unitKerja->nama_unit_kerja, $htmlFlags, 'UTF-8'));
            $templateProcessor->setValue("kode_barang#{$index}", htmlspecialchars($detail->barang->kategori->kode_barang, $htmlFlags, 'UTF-8'));
            $templateProcessor->setValue("nama_barang#{$index}", htmlspecialchars($detail->barang->nama_barang, $htmlFlags, 'UTF-8'));
            $templateProcessor->setValue("spesifikasi_nama_barang#{$index}", htmlspecialchars($detail->barang->spesifikasi_nama_barang, $htmlFlags, 'UTF-8'));
            $templateProcessor->setValue("total_permintaan#{$index}", htmlspecialchars($detail->jumlah_permintaan, $htmlFlags, 'UTF-8'));
            $templateProcessor->setValue("stok_awal#{$index}", htmlspecialchars($stok_awal, $htmlFlags, 'UTF-8'));
            // $templateProcessor->setValue("jumlah#{$index}", $detail->barang->jumlah);
            $templateProcessor->setValue("jumlah#{$index}", htmlspecialchars($sisa_persediaan, $htmlFlags, 'UTF-8'));
            $templateProcessor->setValue("usulan_pengajuan_persetujuan#{$index}", htmlspecialchars($usulan_pengajuan_persetujuan, $htmlFlags, 'UTF-8'));
            $templateProcessor->setValue("satuan#{$index}", htmlspecialchars($detail->barang->satuan, $htmlFlags, 'UTF-8'));
            $templateProcessor->setValue("keperluan#{$index}", htmlspecialchars($detail->permintaan->keperluan, $htmlFlags, 'UTF-8'));
        }

        // save file baru
        $fileName = 'SPB_' . $tanggalCetak . '.docx';
        $tempFilePath = storage_path('app/public/' . $fileName);
        $templateProcessor->saveAs($tempFilePath);
        return response()->download($tempFilePath)->deleteFileAfterSend(true);
    }



    
}
