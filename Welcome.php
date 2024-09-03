<?php
defined('BASEPATH') OR exit('No direct script access allowed');

use PhpOffice\PhpWord\IOFactory;
use Dompdf\Dompdf;
use Mpdf\Mpdf;



class Welcome extends CI_Controller {

	/**
	 * Index Page for this controller.
	 *
	 * Maps to the following URL
	 * 		http://example.com/index.php/welcome
	 *	- or -
	 * 		http://example.com/index.php/welcome/index
	 *	- or -
	 * Since this controller is set as the default controller in
	 * config/routes.php, it's displayed at http://example.com/
	 *
	 * So any other public methods not prefixed with an underscore will
	 * map to /index.php/welcome/<method_name>
	 * @see https://codeigniter.com/userguide3/general/urls.html
	 */
	public function index()
	{
		$this->load->view('welcome_message');
	}

	public function uploadAndConvert()
    {
        // Load helper
        $this->load->helper('url');

        // Cek apakah file di-upload
        if (!empty($_FILES['wordFile']['name'])) {
            $config['upload_path'] = FCPATH . 'uploads/';
            $config['allowed_types'] = 'docx';
            $config['file_name'] = pathinfo($_FILES['wordFile']['name'], PATHINFO_FILENAME);

            // Load library upload dan inisialisasi
            $this->load->library('upload', $config);

            if ($this->upload->do_upload('wordFile')) {
                $data = $this->upload->data();
                $filename = $data['file_name'];
                $filenameWithoutExtension = pathinfo($filename, PATHINFO_FILENAME);

                // Panggil metode konversi
                $this->convertDocxToPdf($filenameWithoutExtension);
            } else {
                echo $this->upload->display_errors();
            }
        } else {
            echo "No file uploaded.";
        }
    }

     public function convertDocxToPdf($filename)
{
    $docxFilePath = FCPATH . 'uploads/' . $filename . '.docx';
    $pdfFilePath = FCPATH . 'uploads/' . $filename . '.pdf';

    $ch = curl_init();

    curl_setopt($ch, CURLOPT_URL, 'https://api.apyhub.com/convert/word-file/pdf-file?output=' . $filename . '.pdf&landscape=false');
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
    curl_setopt($ch, CURLOPT_CUSTOMREQUEST, 'POST');
    curl_setopt($ch, CURLOPT_HTTPHEADER, [
        'apy-token: APY0m7F47eJCYSYRuJYbN9u8ImFT5JcdKcoykeyCHHJqeM0UoQ1SMpNp94ePFcqjioldAa',
        'content-type: multipart/form-data',
    ]);
    curl_setopt($ch, CURLOPT_POSTFIELDS, [
        'file' => new CURLFile($docxFilePath),
    ]);

    $response = curl_exec($ch);

    if (curl_errno($ch)) {
        echo 'Error:' . curl_error($ch);
    } else {
        // Simpan respons (PDF) ke dalam file di folder uploads/
        file_put_contents($pdfFilePath, $response);

        // Mengirimkan file ke browser untuk didownload
        $this->downloadFile($pdfFilePath);
    }

    curl_close($ch);
}

private function downloadFile($filePath)
{
    // Periksa apakah file ada
    if (file_exists($filePath)) {
        // Menyiapkan header untuk unduhan file
        header('Content-Description: File Transfer');
        header('Content-Type: application/pdf');
        header('Content-Disposition: attachment; filename="' . basename($filePath) . '"');
        header('Expires: 0');
        header('Cache-Control: must-revalidate');
        header('Pragma: public');
        header('Content-Length: ' . filesize($filePath));

        // Membaca file dan mengirimkannya ke output
        readfile($filePath);
        exit;
    } else {
        echo "File not found.";
    }
}


//     public function convertDocxToPdf($filename)
//     {
//         require_once FCPATH . 'vendor/autoload.php';

//         $docxFilePath = FCPATH . 'uploads/' . $filename . '.docx';

//         if (!file_exists($docxFilePath)) {
//             die('File DOCX tidak ditemukan!');
//         }

//         $phpWord = \PhpOffice\PhpWord\IOFactory::load($docxFilePath, 'Word2007');
    
//         // Ambil header
//         $headerHtml = '';
//         $headers = $phpWord->getSections()[0]->getHeaders();
//         foreach ($headers as $header) {
//             foreach ($header->getElements() as $element) {
//                 $headerHtml .= $this->elementToHtmlHeader($element);
//             }
//         }

//         $headerHtmlText = [];
//         $headers = $phpWord->getSections()[0]->getHeaders();
//         foreach ($headers as $header) {
//             foreach ($header->getElements() as $element) {
//                 // $headerHtmlText .= $this->elementToHtmlHeaderText($element);
//                 array_push($headerHtmlText, $this->elementToHtmlHeaderText($element));
//             }
//         }

//         // Ambil footer
//         $footerHtml = '';
//         $footers = $phpWord->getSections()[0]->getFooters();
//         $indexFooter = -1;
//         $totalPage = count($footers);
//         foreach ($footers as $footer) {
//             $indexFooter ++;
//             foreach ($footer->getElements() as $element) {
//                 $footerHtml .= $this->elementToHtmlFooter($element,$indexFooter,$totalPage);
//             }
//         }

//         $htmlFilePath = FCPATH . 'uploads/' . $filename . '.html';
//         $htmlWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'HTML');
//         $htmlWriter->save($htmlFilePath);

//         $htmlContent = file_get_contents($htmlFilePath);
//         $imagePath = 'https://word.pundirupiah.my.id/assets/rget.png'; 
//         $imagePathgaruda = 'https://word.pundirupiah.my.id/assets/garuda.png'; 
//         $imagePathqr = 'https://word.pundirupiah.my.id/assets/qr.png'; 
        
//         $htmlContentHeader = '
//         <table border="1" cellpadding="5" cellspacing="0">
//             <tr>
//                 <td rowspan="2" style="text-align: center; vertical-align: middle;"><img src="' . $imagePath . '" width="100" style="display: block;"></td>
//                 <td colspan="4" style="text-align: center; vertical-align: middle;">' . (isset($headerHtmlText[1][1][2]) ? htmlspecialchars($headerHtmlText[1][1][2], ENT_QUOTES, 'UTF-8') : '') . '</td>
//                 <td rowspan="2" style="text-align: center; vertical-align: middle;"><img src="' . $imagePathgaruda . '" width="40" style="display: block;"></td>
//             </tr>
//             <tr>
//                 <td colspan="4" style="text-align: center; vertical-align: middle;">' . (isset($headerHtmlText[1][2][2]) ? htmlspecialchars($headerHtmlText[1][2][2], ENT_QUOTES, 'UTF-8') : '') . '</td>
//             </tr>
//             <tr>
//                 <td> System No.</td>
//                 <td> ' . (isset($headerHtmlText[1][3][1]) ? htmlspecialchars($headerHtmlText[1][3][1], ENT_QUOTES, 'UTF-8') : '') . '</td>
//                 <td> Sub System No.</td>
//                 <td> ' . (isset($headerHtmlText[1][3][3]) ? htmlspecialchars($headerHtmlText[1][3][3], ENT_QUOTES, 'UTF-8') : '') . '</td>
//                 <td> Discipline</td>
//                 <td> Piping</td>
//             </tr>
//             <tr>
//                 <td> Printed Date </td>
//                 <td> ' . (isset($headerHtmlText[1][4][1]) ? htmlspecialchars($headerHtmlText[1][4][1], ENT_QUOTES, 'UTF-8') : '') . ' </td>
//                 <td> Location </td>
//                 <td> opf</td>
//                 <td> Unique No</td>
//                 <td> ' . (isset($headerHtmlText[1][4][5]) ? htmlspecialchars($headerHtmlText[1][4][5], ENT_QUOTES, 'UTF-8') : '') . '</td>
//             </tr>
//             <tr>
//                 <td> Tag No.</td>
//                 <td colspan="5">' . (isset($headerHtmlText[1][5][1]) ? htmlspecialchars($headerHtmlText[1][5][1], ENT_QUOTES, 'UTF-8') : '') . '</td>
//             </tr>
//             <tr>
//                 <td>Sub System Description </td>
//                 <td colspan="4">' . (isset($headerHtmlText[1][6][1]) ? htmlspecialchars($headerHtmlText[1][6][1], ENT_QUOTES, 'UTF-8') : '') . '</td>
//                 <td rowspan="3" style="text-align: center; vertical-align: middle;"> <img src="' . $imagePathqr . '" width="40" style="display: block;"></td>
//             </tr>
//             <tr>
//                 <td> Equipment Description </td>
//                 <td colspan="4">' . (isset($headerHtmlText[1][7][1]) ? htmlspecialchars($headerHtmlText[1][7][1], ENT_QUOTES, 'UTF-8') : '') . '</td>
//             </tr>
//             <tr>
//                 <td>Drawing No.</td>
//                 <td colspan="4">' . (isset($headerHtmlText[1][8][1]) ? htmlspecialchars($headerHtmlText[1][8][1], ENT_QUOTES, 'UTF-8') : '') . '</td>
//             </tr>
//             <tr>
//                 <td>Check Record No.</td>
//                 <td>' . (isset($headerHtmlText[1][9][1]) ? htmlspecialchars($headerHtmlText[1][9][1], ENT_QUOTES, 'UTF-8') : '') . '</td>
//                 <td colspan="4">' . (isset($headerHtmlText[1][9][2]) ? htmlspecialchars($headerHtmlText[1][9][2], ENT_QUOTES, 'UTF-8') : '') . '</td>
//             </tr>
//         </table>';
    
//         $htmlContentFooter = '
//             <table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse; border-color: white;">
//                 <tr>
//                     <td rowspan="2" style="text-align: left; vertical-align: middle; border: 1px solid white;">'.$footerHtml.'</td>
//                     <td rowspan="2" style="text-align: right; vertical-align: middle; border: 1px solid white;">P1 / 1</td>
//                 </tr>
//             </table>';

//         // Template HTML dengan header dan footer untuk setiap halaman
//         $fullHtmlContent = '
//             <html>
//                 <head>
//                     <style>
//                         @page {
//                             margin: 50px 55px;
//                         }
//                         header {
//                             position: fixed;
//                             top: -10px;
//                             left: 0;
//                             right: 0;
//                             height: 50px;
//                             text-align: center;
//                             font-size: 12px;
//                             color: #000;
//                         }
//                         footer {
//                             position: fixed;
//                             bottom: -40px;
//                             left: 0;
//                             right: 0;
//                             height: 50px;
//                             text-align: center;
//                             font-size: 12px;
//                             color: #000;
//                         }
//                         .content {
//                             margin-top: 140px;
//                         }
//                          table {
//                             border-collapse: collapse;
//                         }
//                         table, th, td {
//                             padding:1px;
//                         }
//                         tr {
//                             line-height: 0.8;
//                         }
//                     </style>
//                 </head>
//                 <body>
//                     <header>' . $htmlContentHeader .  '</header>
//                     <footer>' . $htmlContentFooter . '</footer>
//                     <div class="content">' . $htmlContent . '</div>
//                 </body>
//             </html>';

//           // Start output buffering to prevent any accidental output
// ob_start();

// $dompdf = new Dompdf();
// $dompdf->set_option('isRemoteEnabled', true);
// $dompdf->loadHtml($fullHtmlContent);
// $dompdf->setPaper('A4', 'portrait');
// $dompdf->render();

// // Path untuk menyimpan file PDF
// $pdfFilePath = FCPATH . 'uploads/' . $filename . '.pdf';

// // Simpan PDF ke folder uploads
// file_put_contents($pdfFilePath, $dompdf->output());

// // Menghapus file HTML sementara
// unlink($htmlFilePath);

// // Flush the output buffer
// ob_end_clean();

// // Output file PDF langsung untuk diunduh
// $pdfContent = $dompdf->output();

// // Set headers untuk mengunduh file PDF
// header('Content-Type: application/pdf');
// header('Content-Disposition: attachment; filename="' . $filename . '.pdf"');
// header('Content-Transfer-Encoding: binary');
// header('Content-Length: ' . strlen($pdfContent));

// // Output file PDF ke browser
// echo $pdfContent;
//             // Berikan notifikasi bahwa file telah disimpan
//             // echo "PDF berhasil disimpan di " . $pdfFilePath;
//     }

//     private function elementToHtmlHeader($element)
//     {
//         $html = '';
    
//         if (is_object($element)) {
//             if (method_exists($element, 'getText')) {
//                 $text = $element->getText();
//                 // Pastikan $text adalah string
//                 if (is_string($text)) {
//                     $html .= '<p style="margin: 0; padding: 0;">' . htmlspecialchars($text, ENT_QUOTES, 'UTF-8') . '</p>';
//                 } else {
//                     // Jika $text bukan string, tampilkan sebagai teks mentah untuk debugging
//                     $html .= '<p style="margin: 0; padding: 0;">' . htmlspecialchars(print_r($text, true), ENT_QUOTES, 'UTF-8') . '</p>';
//                 }
//             } elseif (method_exists($element, 'getElements')) {
//                 foreach ($element->getElements() as $childElement) {
//                     $html .= $this->elementToHtmlHeader($childElement);
//                 }
//             } elseif (method_exists($element, 'getRows')) {
//                 $html .= '<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse; margin: 0; padding: 0;">';
//                 foreach ($element->getRows() as $row) {
//                     $html .= '<tr>';
//                     foreach ($row->getCells() as $cell) {
//                         $html .= '<td style="padding: 0; margin: 0; border: 0;">' . $this->elementToHtmlHeader($cell) . '</td>';
//                     }
//                     $html .= '</tr>';
//                 }
//                 $html .= '</table>';
//             } elseif (method_exists($element, 'getTextRun')) {
//                 foreach ($element->getTextRun()->getElements() as $textElement) {
//                     $html .= $this->elementToHtmlHeader($textElement);
//                 }
//             } elseif (method_exists($element, 'getImage')) {
//                 $imageSrc = $element->getImageSrc();
//                 $html .= '<p>Image Source Debug: ' . htmlspecialchars($imageSrc, ENT_QUOTES, 'UTF-8') . '</p>';
    
//                 if (is_string($imageSrc)) {
//                     $html .= '<img src="' . htmlspecialchars($imageSrc, ENT_QUOTES, 'UTF-8') . '" alt="Image" style="display: block; margin: 0; padding: 0;" />';
//                 } else {
//                     $html .= '<p>' . htmlspecialchars(print_r($imageSrc, true), ENT_QUOTES, 'UTF-8') . '</p>';
//                 }
//             }
//         } elseif (is_array($element)) {
//             foreach ($element as $subElement) {
//                 $html .= $this->elementToHtmlHeader($subElement);
//             }
//         } else {
//             // Jika elemen bukan objek atau array, anggap sebagai string
//             $html .= '<p style="margin: 0; padding: 0;">' . htmlspecialchars((string)$element, ENT_QUOTES, 'UTF-8') . '</p>';
//         }
    
//         return $html;
//     }

//     private function elementToHtmlHeaderText($element)
// {
//     $html = []; // Array untuk menampung semua teks dari elemen

//     if (is_object($element)) {
//         if (method_exists($element, 'getText')) {
//             $text = $element->getText();
//             if (is_string($text)) {
//                 // Menyimpan teks dari elemen ke dalam array
//                 $html[] = $text;
//             }
//         } elseif (method_exists($element, 'getElements')) {
//             foreach ($element->getElements() as $childElement) {
//                 $html = array_merge($html, $this->elementToHtmlHeaderText($childElement));
//             }
//         } elseif (method_exists($element, 'getRows')) {
//             foreach ($element->getRows() as $row) {
//                 $rowContent = []; // Array untuk menampung teks dari setiap sel dalam satu baris
//                 foreach ($row->getCells() as $cell) {
//                     $cellText = $this->elementToHtmlHeaderText($cell);
//                     $rowContent[] = implode(' ', $cellText); // Gabungkan teks dari setiap sel
//                 }
//                 $html[] = $rowContent; // Tambahkan baris ke array $html
//             }
//         } elseif (method_exists($element, 'getTextRun')) {
//             foreach ($element->getTextRun()->getElements() as $textElement) {
//                 $html = array_merge($html, $this->elementToHtmlHeaderText($textElement));
//             }
//         }
//     } elseif (is_array($element)) {
//         foreach ($element as $subElement) {
//             $html = array_merge($html, $this->elementToHtmlHeaderText($subElement));
//         }
//     } else {
//         $html[] = (string)$element; // Perlakukan elemen lain sebagai string
//     }

//     return $html; // Mengembalikan array teks
// }


    

//     private function elementToHtmlFooter($element, $indexFooter, $totalPage)
//     {
//         $html = '';
    
//         if (is_object($element)) {
//             if (method_exists($element, 'getText')) {
//                 $text = $element->getText();
    
//                 // Cek jika $text adalah array dan gabungkan menjadi string
//                 if (is_array($text)) {
//                     $text = implode(' ', $text);
//                 }
    
//                 // Pastikan $text adalah string sebelum memproses lebih lanjut
//                 if (is_string($text)) {
//                     // Gunakan regex untuk mencari placeholder yang mungkin berbeda formatnya
//                     // $text = preg_replace('/\{ PAGE [^}]*\}/i', $indexFooter, $text);
//                     // $text = preg_replace('/\{ NUMPAGES [^}]*\}/i', $totalPage, $text);
    
//                     // Hapus placeholder lain yang mungkin tersisa
//                     $text = preg_replace('/\{[^}]+\}/', '', $text);
    
//                     // Render HTML untuk teks yang sudah dibersihkan
//                     $html .= htmlspecialchars($text, ENT_QUOTES, 'UTF-8');
//                 }
//             }
//         }
    
//         return $html;
//     }
    
    
      
}
