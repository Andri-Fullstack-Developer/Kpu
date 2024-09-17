<?php

namespace App\Http\Controllers;

use App\Models\User;
use Illuminate\Http\Request;
use PhpOffice\PhpSpreadsheet\IOFactory;

class ImpportDataContoller extends Controller
{
    public function importData(Request $req)
    {
        // Validasi file yang diunggah
        $req->validate([
            'upload' => 'required|mimes:xls,xlsx',
        ]);

        // Ambil file dari request
        $file = $req->file('upload');

        try {
            // Baca file Excel menggunakan PhpSpreadsheet
            $spreadsheet = IOFactory::load($file->getRealPath());
            $sheet = $spreadsheet->getActiveSheet();
            $rows = $sheet->toArray();
            $rowCount = count($rows); // Hitung jumlah baris

            // dd($rowCount);

            // Cek apakah jumlah baris lebih dari 50 atau kurang dari 50
            if ($rowCount > 56) { // Lebih dari 51 karena termasuk header
                return redirect()->back()->with('error', 'The file contains too many rows. Maximum allowed is 55 rows.');
            } elseif ($rowCount < 56) {
                return redirect()->back()->with('error', 'The file contains fewer than 55 rows.');
            }

            // Ambil baris header untuk mapping kolom
            $header = array_map('trim', $rows[0]); // Baris pertama sebagai header
            unset($rows[0]); // Hapus baris header dari data

            // Periksa duplikasi username di file
            $usernames = [];
            $duplicateUsernames = [];

            foreach ($rows as $index => $row) {
                $rowAssoc = array_combine($header, $row);
                $username = $rowAssoc['Username'] ?? null;
                if ($username === null) {
                    continue; // Lewati baris jika username tidak ada
                }
                if (in_array($username, $usernames)) {
                    $duplicateUsernames[] = $username;
                } else {
                    $usernames[] = $username;
                }
            }

            if (!empty($duplicateUsernames)) {
                // Kembalikan error jika ada duplikasi
                return redirect()->back()->with('error', 'File contains duplicate usernames: ' . implode(', ', $duplicateUsernames));
            }

            // Ambil username yang sudah ada
            $existingUsernames = User::pluck('username')->toArray();

            // Proses data jika validasi lolos
            $newRows = [];
            foreach ($rows as $row) {
                $rowAssoc = array_combine($header, $row);

                // Cek apakah username sudah ada
                if (in_array($rowAssoc['Username'], $existingUsernames)) {
                    // Jika username sudah ada, simpan informasi untuk pengembalian
                    $newRows[] = $rowAssoc['Username'];
                    continue;
                }

                // Update atau buat data di tabel Users
                User::updateOrCreate(
                    [
                        'username' => $rowAssoc['Username'], // Kolom C (Username)
                    ],
                    [
                        'name' => $rowAssoc['Name'], // Kolom 'Nama'
                        'kelas' => $rowAssoc['Kelas'], // Kolom 'Kelas'
                        'password' => bcrypt($rowAssoc['Password']), // Kolom 'Password'
                        'status' => 'aktif',
                        'created_at' => now(),
                        'updated_at' => now(),
                    ]
                );
            }

            if (!empty($newRows)) {
                $duplicateUsernames = implode(', ', $newRows);
                return redirect()->back()->with('error', 'The following usernames already exist and were not imported: ' . $duplicateUsernames);
            }

            // Jika berhasil
            return redirect()->back()->with('success', 'Data imported successfully');
        } catch (\Exception $e) {
            // Tangkap jika ada error lain
            return redirect()->back()->with('error', 'There was an error processing the file: ' . $e->getMessage());
        }
    }
}