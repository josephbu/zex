<?php
/**
 * Zip Export EXpress (ZEX)
 * Upload a ZIP of numbered PNGs → get PDF + PPTX download
 *
 * Requirements:
 *   - PHP 7.4+
 *   - php-zip extension
 *   - Composer: composer require tecnickcom/tcpdf
 *   - Composer: composer require phpoffice/phppresentation
 *
 * Quick setup:
 *   composer require tecnickcom/tcpdf phpoffice/phppresentation
 */

// ─── Config ──────────────────────────────────────────────────────────────────
define('UPLOAD_DIR',  __DIR__ . '/tmp/');
define('OUTPUT_DIR',  __DIR__ . '/tmp/');
define('MAX_MB',              250);
/** Web-optimized PDF only: JPEG quality 1–100 (higher = larger files). */
define('PDF_WEB_JPEG_QUALITY', 90);
/** Full-quality PDF when using the native GD PDF path (non–web-optimized). */
define('PDF_JPEG_QUALITY_HIGH', 95);
define('SLIDE_W_PX',  1920);  // Expected PNG width  (adjust if needed)
define('SLIDE_H_PX',  1080);  // Expected PNG height (adjust if needed)
// Presentation dimensions in EMU (English Metric Units): 1 inch = 914400 EMU
// 16:9 → 10" x 5.625"
define('PPTX_W_EMU',  9144000);
define('PPTX_H_EMU',  5143500);

// ─── Appearance ──────────────────────────────────────────────────────────────
// THEME: 'light-breezy' | 'business-clean' | 'dark-moody'
define('UI_THEME', 'light-breezy');
/** Shown in footer: "&copy; {year} {COPYRIGHT_HOLDER}. All rights reserved." */
define('COPYRIGHT_HOLDER', 'JB');

// ─── Helpers ─────────────────────────────────────────────────────────────────
function cleanup(string $dir): void {
    if (!is_dir($dir)) return;
    foreach (glob($dir . '*') as $f) {
        if (is_file($f)) unlink($f);
    }
}

function natural_sort_pngs(array $files): array {
    usort($files, fn($a, $b) => strnatcasecmp(basename($a), basename($b)));
    return $files;
}

function json_error(string $msg, int $code = 400): void {
    http_response_code($code);
    echo json_encode(['error' => $msg]);
    exit;
}

// ─── Cleanup ─────────────────────────────────────────────────────────────────
// Run on every request (lightweight — scans tmp/ for old session dirs).
// Deletes session folders older than this many seconds so tmp/ never accumulates.
function cleanup_old_sessions(int $maxAgeSeconds = 300): void {
    $base = UPLOAD_DIR;
    if (!is_dir($base)) return;
    foreach (glob($base . '*', GLOB_ONLYDIR) as $dir) {
        if ((time() - filemtime($dir)) < $maxAgeSeconds) continue;
        // Delete all files inside, then the dir itself
        $rii = new RecursiveIteratorIterator(
            new RecursiveDirectoryIterator($dir, RecursiveDirectoryIterator::SKIP_DOTS),
            RecursiveIteratorIterator::CHILD_FIRST
        );
        foreach ($rii as $f) {
            $f->isDir() ? rmdir($f->getPathname()) : unlink($f->getPathname());
        }
        @rmdir($dir);
    }
}
cleanup_old_sessions();

// ─── Router ──────────────────────────────────────────────────────────────────
$action = $_GET['action'] ?? '';

if ($action === 'convert') {
    handle_convert();
} elseif ($action === 'download') {
    handle_download();
} else {
    serve_ui();
}

// ─── Convert handler ─────────────────────────────────────────────────────────
function handle_convert(): void {
    header('Content-Type: application/json');

    if (!isset($_FILES['zipfile'])) {
        json_error('No file uploaded.');
    }

    $file = $_FILES['zipfile'];
    if ($file['error'] !== UPLOAD_ERR_OK) {
        json_error('Upload error code: ' . $file['error']);
    }

    $maxBytes = MAX_MB * 1024 * 1024;
    if ($file['size'] > $maxBytes) {
        json_error('File exceeds ' . MAX_MB . ' MB limit.');
    }

    $ext = strtolower(pathinfo($file['name'], PATHINFO_EXTENSION));
    if ($ext !== 'zip') {
        json_error('Please upload a .zip file.');
    }

    // Create unique session dir
    $sessionId = bin2hex(random_bytes(8));
    $workDir   = UPLOAD_DIR . $sessionId . '/';
    mkdir($workDir, 0755, true);

    // Extract ZIP
    $zip = new ZipArchive();
    if ($zip->open($file['tmp_name']) !== true) {
        json_error('Could not open ZIP file.');
    }

    $zip->extractTo($workDir);
    $zip->close();

    // Find PNGs (recursively, skip hidden/system files)
    $pngs = [];
    $rii  = new RecursiveIteratorIterator(new RecursiveDirectoryIterator($workDir));
    foreach ($rii as $f) {
        if ($f->isFile() && strtolower($f->getExtension()) === 'png') {
            $basename = $f->getBasename();
            // Skip macOS metadata files
            if (str_starts_with($basename, '._') || str_starts_with($basename, '__MACOSX')) continue;
            $pngs[] = $f->getPathname();
        }
    }

    if (empty($pngs)) {
        json_error('No PNG files found in the ZIP.');
    }

    $pngs = natural_sort_pngs($pngs);

    // Derive a clean base name from the uploaded ZIP filename.
    // Keep spaces and common punctuation; only strip chars unsafe in filenames/HTTP headers.
    $rawName  = pathinfo($file['name'], PATHINFO_FILENAME);
    $baseName = preg_replace('#[/\\:*?"<>|]#', '', $rawName);
    $baseName = trim($baseName) ?: 'slides';

    $wantPdf  = isset($_POST['out_pdf']) && $_POST['out_pdf'] === '1';
    $wantPptx = isset($_POST['out_pptx']) && $_POST['out_pptx'] === '1';
    if (!$wantPdf && !$wantPptx) {
        json_error('Select at least one output format (PDF or PowerPoint).');
    }

    $pdfWeb = $wantPdf && isset($_POST['pdf_web']) && $_POST['pdf_web'] === '1';

    // Generate files (only requested formats)
    $pdfPath  = $workDir . $baseName . '.pdf';
    $pptxPath = $workDir . $baseName . '.pptx';

    $pdfOk  = $wantPdf ? generate_pdf($pngs, $pdfPath, $pdfWeb) : false;
    $pptxOk = $wantPptx ? generate_pptx($pngs, $pptxPath) : false;

    $results = [];
    if ($wantPdf && $pdfOk) {
        $results['pdf']     = $sessionId . '/' . $baseName . '.pdf';
        $results['pdfName'] = $baseName . '.pdf';
    }
    if ($wantPptx && $pptxOk) {
        $results['pptx']     = $sessionId . '/' . $baseName . '.pptx';
        $results['pptxName'] = $baseName . '.pptx';
    }

    if (empty($results)) {
        json_error('File generation failed. Check server logs.', 500);
    }

    $results['slideCount'] = count($pngs);
    echo json_encode($results);
}

/** Encode PNG as JPEG in memory (same pixel dimensions). Flattens transparency onto white. */
function png_to_jpeg_binary(string $pngPath, int $quality): ?string {
    if (!function_exists('imagecreatefrompng')) {
        return null;
    }
    $im = @imagecreatefrompng($pngPath);
    if ($im === false) {
        return null;
    }
    $w = imagesx($im);
    $h = imagesy($im);
    $bg = imagecreatetruecolor($w, $h);
    if ($bg === false) {
        imagedestroy($im);
        return null;
    }
    $white = imagecolorallocate($bg, 255, 255, 255);
    imagefill($bg, 0, 0, $white);
    imagealphablending($bg, true);
    imagecopy($bg, $im, 0, 0, 0, 0, $w, $h);
    imagedestroy($im);
    $q = max(0, min(100, $quality));
    ob_start();
    imagejpeg($bg, null, $q);
    $data = ob_get_clean();
    imagedestroy($bg);
    return ($data !== false && strlen($data) > 0) ? $data : null;
}

// ─── PDF generation ──────────────────────────────────────────────────────────
function generate_pdf(array $pngs, string $outPath, bool $webOptimized = false): bool {
    // Check for composer autoload
    $autoload = __DIR__ . '/vendor/autoload.php';
    if (!file_exists($autoload)) {
        // Fallback: use Imagick if available
        if (extension_loaded('imagick')) {
            return generate_pdf_imagick($pngs, $outPath, $webOptimized);
        }
        return generate_pdf_native($pngs, $outPath, $webOptimized);
    }

    require_once $autoload;

    if (!class_exists('TCPDF')) {
        return generate_pdf_native($pngs, $outPath, $webOptimized);
    }

    try {
        // Get dimensions from first image
        [$w, $h] = getimagesize($pngs[0]) ?: [SLIDE_W_PX, SLIDE_H_PX];
        $ratio    = $w / $h;

        // Use mm units. Standard widescreen: 254mm × 142.875mm
        $pageW = 254;
        $pageH = round($pageW / $ratio, 3);

        $pdf = new TCPDF('L', 'mm', [$pageW, $pageH], true, 'UTF-8', false);
        $pdf->SetCreator('ZEX - Zip Export EXpress');
        $pdf->SetAuthor('ZEX - Zip Export EXpress');
        $pdf->SetTitle('Slides');
        $pdf->SetPrintHeader(false);
        $pdf->SetPrintFooter(false);
        $pdf->SetMargins(0, 0, 0);
        $pdf->SetAutoPageBreak(false, 0);

        foreach ($pngs as $png) {
            $pdf->AddPage();
            if ($webOptimized) {
                $jpg = png_to_jpeg_binary($png, PDF_WEB_JPEG_QUALITY);
                if ($jpg !== null) {
                    $pdf->Image('@' . $jpg, 0, 0, $pageW, $pageH, 'JPEG', '', '', false, 300);
                } else {
                    $pdf->Image($png, 0, 0, $pageW, $pageH, 'PNG', '', '', false, 300);
                }
            } else {
                $pdf->Image($png, 0, 0, $pageW, $pageH, 'PNG', '', '', false, 300);
            }
        }

        $pdf->Output($outPath, 'F');
        return file_exists($outPath);
    } catch (Exception $e) {
        error_log('TCPDF error: ' . $e->getMessage());
        return false;
    }
}

function generate_pdf_imagick(array $pngs, string $outPath, bool $webOptimized = false): bool {
    try {
        $im = new Imagick();
        foreach ($pngs as $png) {
            $page = new Imagick($png);
            if ($webOptimized) {
                $page->setImageBackgroundColor(new ImagickPixel('white'));
                if ($page->getImageAlphaChannel()) {
                    $page->setImageAlphaChannel(Imagick::ALPHACHANNEL_REMOVE);
                    $page = $page->mergeImageLayers(Imagick::LAYERMETHOD_FLATTEN);
                }
                $page->setImageCompression(Imagick::COMPRESSION_JPEG);
                $page->setImageCompressionQuality(PDF_WEB_JPEG_QUALITY);
            }
            $page->setImageFormat('pdf');
            $im->addImage($page);
        }
        $im->setFormat('pdf');
        $im->writeImages($outPath, true);
        return file_exists($outPath);
    } catch (Exception $e) {
        error_log('Imagick error: ' . $e->getMessage());
        return false;
    }
}

// Pure PHP fallback using fpdf-style raw PDF construction
function generate_pdf_native(array $pngs, string $outPath, bool $webOptimized = false): bool {
    // Minimal PDF using GD + raw PDF construction
    // This basic version embeds each PNG as a page using raw PDF spec
    try {
        $objects = [];
        $pageIds = [];
        $objNum  = 0;

        $addObj = function(string $content) use (&$objects, &$objNum): int {
            $objNum++;
            $objects[$objNum] = $content;
            return $objNum;
        };

        // Catalog & pages placeholder (will fill page IDs later)
        $catalogId = $addObj(''); // placeholder
        $pagesId   = $addObj(''); // placeholder

        // DPI for PDF: use 96dpi → 1px = 1/96 inch = 72/96 pt
        $dpi = 96;
        $pxToPt = 72 / $dpi;

        $jpegQ = $webOptimized ? PDF_WEB_JPEG_QUALITY : PDF_JPEG_QUALITY_HIGH;

        foreach ($pngs as $png) {
            [$imgW, $imgH] = getimagesize($png);

            $ptW = round($imgW * $pxToPt, 2);
            $ptH = round($imgH * $pxToPt, 2);

            $jpgData = png_to_jpeg_binary($png, $jpegQ);
            if ($jpgData === null) {
                return false;
            }

            $imgId = $addObj("<</Type /XObject /Subtype /Image /Width $imgW /Height $imgH "
                . "/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode "
                . "/Length " . strlen($jpgData) . ">>\nstream\n" . $jpgData . "\nendstream");

            $resId      = $addObj("<</ProcSet [/PDF /ImageC] /XObject <</Img $imgId 0 R>>>>");
            $contentStr = "q $ptW 0 0 $ptH 0 0 cm /Img Do Q";
            $contentId  = $addObj("<</Length " . strlen($contentStr) . ">>\nstream\n$contentStr\nendstream");
            $pageId     = $addObj("<</Type /Page /Parent $pagesId 0 R "
                . "/MediaBox [0 0 $ptW $ptH] "
                . "/Resources $resId 0 R "
                . "/Contents $contentId 0 R>>");
            $pageIds[] = $pageId;
        }

        // Fill pages object
        $kidsStr = implode(' 0 R ', $pageIds) . ' 0 R';
        $objects[$pagesId]  = "<</Type /Pages /Kids [$kidsStr] /Count " . count($pageIds) . ">>";
        $objects[$catalogId] = "<</Type /Catalog /Pages $pagesId 0 R>>";

        // Build PDF binary
        $body    = "%PDF-1.4\n";
        $offsets = [];
        foreach ($objects as $id => $content) {
            $offsets[$id] = strlen($body);
            $body .= "$id 0 obj\n$content\nendobj\n";
        }

        $xrefOffset = strlen($body);
        $count      = count($objects) + 1;
        $body .= "xref\n0 $count\n";
        $body .= "0000000000 65535 f \n";
        foreach ($offsets as $offset) {
            $body .= str_pad($offset, 10, '0', STR_PAD_LEFT) . " 00000 n \n";
        }
        $body .= "trailer\n<</Size $count /Root $catalogId 0 R>>\n";
        $body .= "startxref\n$xrefOffset\n%%EOF";

        file_put_contents($outPath, $body);
        return file_exists($outPath);
    } catch (Exception $e) {
        error_log('Native PDF error: ' . $e->getMessage());
        return false;
    }
}

// ─── PPTX generation ─────────────────────────────────────────────────────────
function generate_pptx(array $pngs, string $outPath): bool {
    $autoload = __DIR__ . '/vendor/autoload.php';
    if (!file_exists($autoload)) {
        return generate_pptx_manual($pngs, $outPath);
    }

    require_once $autoload;

    if (!class_exists('PhpOffice\PhpPresentation\PhpPresentation')) {
        return generate_pptx_manual($pngs, $outPath);
    }

    try {
        use_pptx_library($pngs, $outPath);
        return file_exists($outPath);
    } catch (Exception $e) {
        error_log('PhpPresentation error: ' . $e->getMessage());
        return generate_pptx_manual($pngs, $outPath);
    }
}

function use_pptx_library(array $pngs, string $outPath): void {
    $prs = new \PhpOffice\PhpPresentation\PhpPresentation();
    $prs->getLayout()->setDocumentLayout(\PhpOffice\PhpPresentation\DocumentLayout::LAYOUT_CUSTOM, true);
    $prs->getLayout()->setCX(PPTX_W_EMU, \PhpOffice\PhpPresentation\DocumentLayout::UNIT_EMU);
    $prs->getLayout()->setCY(PPTX_H_EMU, \PhpOffice\PhpPresentation\DocumentLayout::UNIT_EMU);

    $first = true;
    foreach ($pngs as $png) {
        if ($first) {
            $slide = $prs->getActiveSlide();
            $first = false;
        } else {
            $slide = $prs->createSlide();
        }

        $shape = new \PhpOffice\PhpPresentation\Shape\Drawing\File();
        $shape->setPath($png);
        $shape->setWidth(intval(PPTX_W_EMU / 9144));  // convert EMU to points-ish
        $shape->setHeight(intval(PPTX_H_EMU / 9144));
        $shape->setOffsetX(0)->setOffsetY(0);
        $slide->addShape($shape);
    }

    $writer = \PhpOffice\PhpPresentation\IOFactory::createWriter($prs, 'PowerPoint2007');
    $writer->save($outPath);
}

// Manual PPTX builder — validated against python-pptx reference output.
// Structure proven correct: includes theme, correct rel ID ordering (layout=rId1, image=rId2).
function generate_pptx_manual(array $pngs, string $outPath): bool {
    try {
        $n   = count($pngs);
        $cx  = PPTX_W_EMU;
        $cy  = PPTX_H_EMU;
        $PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
        $R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
        $P   = 'http://schemas.openxmlformats.org/presentationml/2006/main';
        $A   = 'http://schemas.openxmlformats.org/drawingml/2006/main';
        $CT  = 'http://schemas.openxmlformats.org/package/2006/content-types';

        $zip = new ZipArchive();
        if ($zip->open($outPath, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) {
            throw new RuntimeException("Cannot create ZIP at $outPath");
        }

        // ── [Content_Types].xml ──────────────────────────────────────────────
        $ct = '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>' . "\n"
            . '<Types xmlns="' . $CT . '">'
            . '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            . '<Default Extension="xml" ContentType="application/xml"/>'
            . '<Default Extension="png" ContentType="image/png"/>'
            . '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
            . '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
            . '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
            . '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>';
        for ($i = 1; $i <= $n; $i++) {
            $ct .= '<Override PartName="/ppt/slides/slide' . $i . '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>';
        }
        $ct .= '</Types>';
        $zip->addFromString('[Content_Types].xml', $ct);

        // ── _rels/.rels ──────────────────────────────────────────────────────
        $zip->addFromString('_rels/.rels',
            '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>' . "\n"
            . '<Relationships xmlns="' . $PKG . '">'
            . '<Relationship Id="rId1" Type="' . $R . '/officeDocument" Target="ppt/presentation.xml"/>'
            . '</Relationships>');

        // ── ppt/_rels/presentation.xml.rels ─────────────────────────────────
        $prRels = '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>' . "\n"
                . '<Relationships xmlns="' . $PKG . '">'
                . '<Relationship Id="rId1" Type="' . $R . '/slideMaster" Target="slideMasters/slideMaster1.xml"/>'
                . '<Relationship Id="rId2" Type="' . $R . '/theme" Target="theme/theme1.xml"/>';
        for ($i = 1; $i <= $n; $i++) {
            $prRels .= '<Relationship Id="rSlide' . $i . '" Type="' . $R . '/slide" Target="slides/slide' . $i . '.xml"/>';
        }
        $prRels .= '</Relationships>';
        $zip->addFromString('ppt/_rels/presentation.xml.rels', $prRels);

        // ── ppt/presentation.xml ─────────────────────────────────────────────
        $sldIds = '';
        for ($i = 1; $i <= $n; $i++) {
            $sldIds .= '<p:sldId id="' . (256 + $i) . '" r:id="rSlide' . $i . '"/>';
        }
        $zip->addFromString('ppt/presentation.xml',
            '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>' . "\n"
            . '<p:presentation xmlns:a="' . $A . '" xmlns:r="' . $R . '" xmlns:p="' . $P . '" saveSubsetFonts="1">'
            . '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
            . '<p:sldIdLst>' . $sldIds . '</p:sldIdLst>'
            . '<p:sldSz cx="' . $cx . '" cy="' . $cy . '" type="custom"/>'
            . '<p:notesSz cx="6858000" cy="9144000"/>'
            . '</p:presentation>');

        // ── ppt/theme/theme1.xml (minimal but required) ──────────────────────
        $zip->addFromString('ppt/theme/theme1.xml',
            '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>' . "\n"
            . '<a:theme xmlns:a="' . $A . '" name="Office Theme">'
            . '<a:themeElements>'
            . '<a:clrScheme name="Office">'
            . '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
            . '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
            . '<a:dk2><a:srgbClr val="1F497D"/></a:dk2>'
            . '<a:lt2><a:srgbClr val="EEECE1"/></a:lt2>'
            . '<a:accent1><a:srgbClr val="4F81BD"/></a:accent1>'
            . '<a:accent2><a:srgbClr val="C0504D"/></a:accent2>'
            . '<a:accent3><a:srgbClr val="9BBB59"/></a:accent3>'
            . '<a:accent4><a:srgbClr val="8064A2"/></a:accent4>'
            . '<a:accent5><a:srgbClr val="4BACC6"/></a:accent5>'
            . '<a:accent6><a:srgbClr val="F79646"/></a:accent6>'
            . '<a:hlink><a:srgbClr val="0000FF"/></a:hlink>'
            . '<a:folHlink><a:srgbClr val="800080"/></a:folHlink>'
            . '</a:clrScheme>'
            . '<a:fontScheme name="Office">'
            . '<a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
            . '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>'
            . '</a:fontScheme>'
            . '<a:fmtScheme name="Office">'
            . '<a:fillStyleLst>'
            . '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            . '<a:gradFill rotWithShape="1"><a:gsLst>'
            . '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>'
            . '<a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs>'
            . '</a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill>'
            . '<a:gradFill rotWithShape="1"><a:gsLst>'
            . '<a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="100000"/><a:satMod val="130000"/></a:schemeClr></a:gs>'
            . '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="100000"/><a:satMod val="350000"/></a:schemeClr></a:gs>'
            . '</a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill>'
            . '</a:fillStyleLst>'
            . '<a:lnStyleLst>'
            . '<a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
            . '<a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
            . '<a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
            . '</a:lnStyleLst>'
            . '<a:effectStyleLst>'
            . '<a:effectStyle><a:effectLst/></a:effectStyle>'
            . '<a:effectStyle><a:effectLst/></a:effectStyle>'
            . '<a:effectStyle><a:effectLst/></a:effectStyle>'
            . '</a:effectStyleLst>'
            . '<a:bgFillStyleLst>'
            . '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            . '<a:gradFill rotWithShape="1"><a:gsLst>'
            . '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs>'
            . '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs>'
            . '</a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill>'
            . '<a:gradFill rotWithShape="1"><a:gsLst>'
            . '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs>'
            . '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs>'
            . '</a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>'
            . '</a:bgFillStyleLst>'
            . '</a:fmtScheme>'
            . '</a:themeElements>'
            . '<a:objectDefaults/>'
            . '<a:extraClrSchemeLst/>'
            . '</a:theme>');

        // ── Slide master ─────────────────────────────────────────────────────
        $zip->addFromString('ppt/slideMasters/_rels/slideMaster1.xml.rels',
            '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>' . "\n"
            . '<Relationships xmlns="' . $PKG . '">'
            . '<Relationship Id="rId1" Type="' . $R . '/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
            . '<Relationship Id="rId2" Type="' . $R . '/theme" Target="../theme/theme1.xml"/>'
            . '</Relationships>');

        $zip->addFromString('ppt/slideMasters/slideMaster1.xml',
            '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>' . "\n"
            . '<p:sldMaster xmlns:a="' . $A . '" xmlns:r="' . $R . '" xmlns:p="' . $P . '">'
            . '<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>'
            . '<p:spTree>'
            . '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            . '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            . '</p:spTree></p:cSld>'
            . '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
            . '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
            . '<p:txStyles><p:titleStyle><a:lstStyle/></p:titleStyle><p:bodyStyle><a:lstStyle/></p:bodyStyle><p:otherStyle><a:lstStyle/></p:otherStyle></p:txStyles>'
            . '</p:sldMaster>');

        // ── Slide layout (blank) ─────────────────────────────────────────────
        $zip->addFromString('ppt/slideLayouts/_rels/slideLayout1.xml.rels',
            '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>' . "\n"
            . '<Relationships xmlns="' . $PKG . '">'
            . '<Relationship Id="rId1" Type="' . $R . '/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
            . '</Relationships>');

        $zip->addFromString('ppt/slideLayouts/slideLayout1.xml',
            '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>' . "\n"
            . '<p:sldLayout xmlns:a="' . $A . '" xmlns:r="' . $R . '" xmlns:p="' . $P . '" type="blank" preserve="1">'
            . '<p:cSld name="Blank"><p:spTree>'
            . '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            . '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            . '</p:spTree></p:cSld>'
            . '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
            . '</p:sldLayout>');

        // ── Slides ───────────────────────────────────────────────────────────
        foreach ($pngs as $idx => $png) {
            $sn  = $idx + 1;
            $img = 'image' . $sn . '.png';

            $imgBytes = file_get_contents($png);
            if ($imgBytes === false) {
                throw new RuntimeException("Cannot read: $png");
            }
            $zip->addFromString('ppt/media/' . $img, $imgBytes);

            // CRITICAL: layout rel is rId1, image rel is rId2 — blip embed must say rId2
            $zip->addFromString('ppt/slides/_rels/slide' . $sn . '.xml.rels',
                '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>' . "\n"
                . '<Relationships xmlns="' . $PKG . '">'
                . '<Relationship Id="rId1" Type="' . $R . '/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
                . '<Relationship Id="rId2" Type="' . $R . '/image" Target="../media/' . $img . '"/>'
                . '</Relationships>');

            $zip->addFromString('ppt/slides/slide' . $sn . '.xml',
                '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>' . "\n"
                . '<p:sld xmlns:a="' . $A . '" xmlns:r="' . $R . '" xmlns:p="' . $P . '">'
                . '<p:cSld><p:spTree>'
                . '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
                . '<p:grpSpPr/>'
                . '<p:pic>'
                . '<p:nvPicPr>'
                . '<p:cNvPr id="2" name="Picture ' . $sn . '"/>'
                . '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
                . '<p:nvPr/>'
                . '</p:nvPicPr>'
                . '<p:blipFill>'
                . '<a:blip r:embed="rId2"/>'
                . '<a:stretch><a:fillRect/></a:stretch>'
                . '</p:blipFill>'
                . '<p:spPr>'
                . '<a:xfrm><a:off x="0" y="0"/><a:ext cx="' . $cx . '" cy="' . $cy . '"/></a:xfrm>'
                . '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
                . '</p:spPr>'
                . '</p:pic>'
                . '</p:spTree></p:cSld>'
                . '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
                . '</p:sld>');
        }

        $zip->close();
        return file_exists($outPath);

    } catch (Exception $e) {
        error_log('Manual PPTX error: ' . $e->getMessage());
        return false;
    }
}

// ─── Download handler ─────────────────────────────────────────────────────────
function handle_download(): void {
    $file = $_GET['file'] ?? '';

    // Security: session ID is 16 hex chars; filename is URL-encoded safe chars + .pdf/.pptx
    $file = rawurldecode($file);
    if (!preg_match('/^[a-f0-9]{16}\/([^\/\\:*?"<>|]+\.(pdf|pptx))$/', $file, $m)) {
        http_response_code(400);
        exit('Invalid file.');
    }

    $path = UPLOAD_DIR . $file;
    if (!file_exists($path)) {
        http_response_code(404);
        exit('File not found.');
    }

    $ext      = $m[2];
    // Use &name= param if provided, otherwise derive from file path
    $filename = isset($_GET['name']) ? basename(rawurldecode($_GET['name'])) : $m[1];
    // Ensure extension is correct regardless of name param
    $filename = preg_replace('/\.' . preg_quote($ext, '/') . '$/i', '', $filename) . '.' . $ext;
    $mime     = $ext === 'pdf'
        ? 'application/pdf'
        : 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
    header('Content-Type: ' . $mime);
    // RFC 5987 encoding handles spaces and special chars safely across all browsers
    $encodedName = rawurlencode($filename);
    header('Content-Disposition: attachment; filename="' . $filename . '"; filename*=UTF-8\'\'' . $encodedName);
    header('Content-Length: ' . filesize($path));
    header('Cache-Control: no-cache');
    readfile($path);
    exit;
}

// ─── UI ───────────────────────────────────────────────────────────────────────
function get_theme_vars(string $theme): array {
    $themes = [
        'light-breezy' => [
            'font_url'    => 'https://fonts.googleapis.com/css2?family=Nunito:wght@300;400;600;700&family=Playfair+Display:ital@0;1&display=swap',
            'title_font'  => "'Playfair Display', Georgia, serif",
            'body_font'   => "'Nunito', sans-serif",
            'bg'          => '#f0f4f8',
            'surface'     => '#ffffff',
            'border'      => '#d1dbe8',
            'accent'      => '#3b9eff',
            'accent2'     => '#7ec8ff',
            'accent_hover'=> '#1a85e8',
            'accent_glow' => 'rgba(59,158,255,0.3)',
            'text'        => '#1a2b3c',
            'muted'       => '#6b8099',
            'success'     => '#10b981',
            'error'       => '#ef4444',
            'pdf_clr'     => '#f97316',
            'ppt_clr'     => '#3b9eff',
            'drop_hover'  => 'rgba(59,158,255,0.06)',
            'tag_line'    => 'Turn your ZIP export of PNGs into PDF or PowerPoint slides!',
            'body_bg'     => 'linear-gradient(160deg, #e8f0fb 0%, #f5f8ff 60%, #eef6f0 100%)',
            'card_shadow' => '0 8px 40px rgba(59,100,180,0.10)',
            'title_grad'  => 'linear-gradient(135deg, #1a2b3c 0%, #3b9eff 100%)',
        ],
        'business-clean' => [
            'font_url'    => 'https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&family=Fraunces:ital,wght@0,400;1,400&display=swap',
            'title_font'  => "'Fraunces', Georgia, serif",
            'body_font'   => "'Inter', sans-serif",
            'bg'          => '#f7f7f5',
            'surface'     => '#ffffff',
            'border'      => '#e0ddd8',
            'accent'      => '#1a1a1a',
            'accent2'     => '#555550',
            'accent_hover'=> '#333330',
            'accent_glow' => 'rgba(26,26,26,0.15)',
            'text'        => '#1a1a1a',
            'muted'       => '#888880',
            'success'     => '#2d6a4f',
            'error'       => '#c0392b',
            'pdf_clr'     => '#c0392b',
            'ppt_clr'     => '#1a1a1a',
            'drop_hover'  => 'rgba(26,26,26,0.03)',
            'tag_line'    => 'Turn your ZIP export of PNGs into PDF or PowerPoint slides!',
            'body_bg'     => '#f7f7f5',
            'card_shadow' => '0 2px 20px rgba(0,0,0,0.07)',
            'title_grad'  => 'none',
        ],
        'dark-moody' => [
            'font_url'    => 'https://fonts.googleapis.com/css2?family=Instrument+Serif:ital@0;1&family=DM+Sans:wght@300;400;500;600&display=swap',
            'title_font'  => "'Instrument Serif', Georgia, serif",
            'body_font'   => "'DM Sans', sans-serif",
            'bg'          => '#0f0f11',
            'surface'     => '#1a1a1e',
            'border'      => '#2e2e35',
            'accent'      => '#6c63ff',
            'accent2'     => '#a78bfa',
            'accent_hover'=> '#7c74ff',
            'accent_glow' => 'rgba(108,99,255,0.35)',
            'text'        => '#e8e8f0',
            'muted'       => '#888899',
            'success'     => '#34d399',
            'error'       => '#f87171',
            'pdf_clr'     => '#fb923c',
            'ppt_clr'     => '#60a5fa',
            'drop_hover'  => 'rgba(108,99,255,0.05)',
            'tag_line'    => 'Turn your ZIP export of PNGs into PDF or PowerPoint slides!',
            'body_bg'     => '#0f0f11',
            'card_shadow' => 'none',
            'title_grad'  => 'linear-gradient(135deg, #e8e8f0 0%, #a78bfa 100%)',
        ],
    ];
    return $themes[$theme] ?? $themes['dark-moody'];
}

function serve_ui(): void {
    // Optional: composer require tecnickcom/tcpdf phpoffice/phppresentation for higher-quality output
    $t = get_theme_vars(UI_THEME);

    $is_light     = (UI_THEME !== 'dark-moody');
    $title_style  = $t['title_grad'] === 'none'
        ? "color: {$t['text']};"
        : "background: {$t['title_grad']}; -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;";
    $card_top_bar = $is_light
        ? "background: linear-gradient(90deg, transparent, {$t['accent']}, transparent); opacity: 0.2;"
        : "background: linear-gradient(90deg, transparent, {$t['accent']}, transparent); opacity: 0.6;";

    echo '<!DOCTYPE html>' . "\n";
    echo '<html lang="en">' . "\n";
    echo '<head>' . "\n";
    echo '<meta charset="UTF-8">' . "\n";
    echo '<meta name="viewport" content="width=device-width, initial-scale=1.0">' . "\n";
    echo '<title>ZEX &ndash; Zip Export EXpress</title>' . "\n";
    echo '<link rel="preconnect" href="https://fonts.googleapis.com">' . "\n";
    echo '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>' . "\n";
    echo '<link href="' . htmlspecialchars($t['font_url']) . '" rel="stylesheet">' . "\n";
    echo '<style>' . "\n";
    echo '  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }' . "\n";
    echo '
  :root {
    --bg:           ' . $t['bg'] . ';
    --surface:      ' . $t['surface'] . ';
    --border:       ' . $t['border'] . ';
    --accent:       ' . $t['accent'] . ';
    --accent2:      ' . $t['accent2'] . ';
    --accent-hover: ' . $t['accent_hover'] . ';
    --accent-glow:  ' . $t['accent_glow'] . ';
    --text:         ' . $t['text'] . ';
    --muted:        ' . $t['muted'] . ';
    --success:      ' . $t['success'] . ';
    --error:        ' . $t['error'] . ';
    --pdf-clr:      ' . $t['pdf_clr'] . ';
    --ppt-clr:      ' . $t['ppt_clr'] . ';
    --drop-hover:   ' . $t['drop_hover'] . ';
    --body-bg:      ' . $t['body_bg'] . ';
    --card-shadow:  ' . $t['card_shadow'] . ';
    --title-font:   ' . $t['title_font'] . ';
    --body-font:    ' . $t['body_font'] . ';
  }

  body {
    font-family: var(--body-font);
    background: var(--body-bg);
    color: var(--text);
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 1.75rem 1.25rem 3.5rem;
  }

  .header {
    text-align: center;
    margin-bottom: 1.75rem;
  }

  .header h1 {
    font-family: var(--body-font);
    font-size: clamp(2rem, 5vw, 3.2rem);
    font-style: normal;
    font-weight: 800;
    letter-spacing: -0.05em;
    line-height: 1.15;
    margin-bottom: 0.5rem;
    ' . $title_style . '
  }

  .header p {
    color: var(--muted);
    font-size: 0.92rem;
    font-weight: 300;
    letter-spacing: 0.01em;
  }

  .app-sub {
    font-family: var(--body-font);
    font-size: 0.78rem;
    font-weight: 600;
    letter-spacing: 0.18em;
    text-transform: uppercase;
    color: var(--muted);
    margin-bottom: 0.4rem;
    opacity: 0.7;
  }

  .x-mark {
    color: var(--accent);
    -webkit-text-fill-color: var(--accent);
  }

  .card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 20px;
    padding: 2rem;
    width: 100%;
    max-width: 560px;
    position: relative;
    overflow: hidden;
    box-shadow: var(--card-shadow);
  }

  .card::before {
    content: "";
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 1px;
    ' . $card_top_bar . '
  }

  .drop-zone {
    border: 2px dashed var(--border);
    border-radius: 14px;
    padding: 2.25rem 1.5rem;
    text-align: center;
    cursor: pointer;
    transition: all 0.25s ease;
    position: relative;
    background: transparent;
  }

  .drop-zone:hover,
  .drop-zone.dragover {
    border-color: var(--accent);
    background: var(--drop-hover);
  }

  .drop-zone input[type=file] {
    position: absolute;
    inset: 0;
    opacity: 0;
    cursor: pointer;
    width: 100%;
    height: 100%;
  }

  .drop-icon {
    font-size: 2.8rem;
    margin-bottom: 0.8rem;
    display: block;
    filter: grayscale(0.3);
  }

  .drop-zone h3 {
    font-size: 1rem;
    font-weight: 500;
    margin-bottom: 0.35rem;
  }

  .drop-zone p {
    color: var(--muted);
    font-size: 0.83rem;
  }

  .file-selected {
    display: none;
    align-items: center;
    gap: 0.8rem;
    background: rgba(52, 211, 153, 0.08);
    border: 1px solid rgba(52, 211, 153, 0.2);
    border-radius: 10px;
    padding: 0.9rem 1.1rem;
    margin-top: 1.2rem;
    font-size: 0.88rem;
  }

  .file-selected.show { display: flex; }
  .file-selected .icon { font-size: 1.4rem; }
  .file-selected .name { font-weight: 500; }
  .file-selected .size { color: var(--muted); font-size: 0.8rem; }

  .output-options {
    margin-top: 1.25rem;
    padding: 1rem 1.1rem;
    border-radius: 12px;
    border: 1px solid var(--border);
    background: var(--drop-hover);
  }

  .output-options-label {
    display: block;
    font-size: 0.78rem;
    font-weight: 600;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: var(--muted);
    margin-bottom: 0.75rem;
  }

  .output-options-row {
    display: flex;
    flex-wrap: wrap;
    gap: 1rem 1.5rem;
  }

  .output-option {
    display: flex;
    align-items: center;
    gap: 0.5rem;
    font-size: 0.9rem;
    cursor: pointer;
    user-select: none;
  }

  .output-option input {
    width: 1.05rem;
    height: 1.05rem;
    accent-color: var(--accent);
    cursor: pointer;
  }

  .output-option input:disabled {
    cursor: not-allowed;
    opacity: 0.45;
  }

  .output-options-sub {
    margin-top: 0.65rem;
    padding-top: 0.65rem;
    border-top: 1px solid var(--border);
  }

  .output-option-supplement {
    font-size: 0.84rem;
    color: var(--muted);
  }

  .btn-convert {
    display: block;
    width: 100%;
    margin-top: 1.5rem;
    padding: 1rem;
    background: var(--accent);
    color: #fff;
    border: none;
    border-radius: 12px;
    font-family: var(--body-font);
    font-size: 1rem;
    font-weight: 600;
    cursor: pointer;
    transition: all 0.2s ease;
    letter-spacing: 0.01em;
  }

  .btn-convert:hover:not(:disabled) {
    background: var(--accent-hover);
    transform: translateY(-1px);
    box-shadow: 0 8px 30px var(--accent-glow);
  }

  .btn-convert:disabled {
    opacity: 0.45;
    cursor: not-allowed;
    transform: none;
  }

  .progress-wrap {
    display: none;
    margin-top: 1.5rem;
    flex-direction: column;
    gap: 0.6rem;
  }

  .progress-wrap.show { display: flex; }

  .progress-bar-bg {
    height: 6px;
    background: var(--border);
    border-radius: 99px;
    overflow: hidden;
  }

  .progress-bar-fill {
    height: 100%;
    background: linear-gradient(90deg, var(--accent), var(--accent2));
    border-radius: 99px;
    width: 0%;
    transition: width 0.4s ease;
  }

  .progress-label {
    font-size: 0.83rem;
    color: var(--muted);
    text-align: center;
  }

  .results {
    display: none;
    margin-top: 1.8rem;
    flex-direction: column;
    gap: 1rem;
  }

  .results.show { display: flex; }

  .results-header {
    text-align: center;
    font-size: 0.88rem;
    color: var(--muted);
  }

  .results-header strong { color: var(--success); }

  .download-btns {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 0.8rem;
  }

  .dl-btn {
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 0.4rem;
    padding: 1.1rem 1rem;
    border-radius: 12px;
    text-decoration: none;
    font-weight: 600;
    font-size: 0.95rem;
    transition: all 0.2s ease;
    border: 1.5px solid transparent;
  }

  .dl-btn .dl-icon { font-size: 1.8rem; }
  .dl-btn .dl-type { font-size: 0.75rem; font-weight: 400; opacity: 0.75; }

  .dl-btn.pdf {
    background: rgba(251, 146, 60, 0.1);
    border-color: rgba(251, 146, 60, 0.3);
    color: var(--pdf-clr);
  }

  .dl-btn.pdf:hover {
    background: rgba(251, 146, 60, 0.18);
    border-color: var(--pdf-clr);
    transform: translateY(-2px);
    box-shadow: 0 6px 24px rgba(251, 146, 60, 0.2);
  }

  .dl-btn.pptx {
    background: rgba(96, 165, 250, 0.1);
    border-color: rgba(96, 165, 250, 0.3);
    color: var(--ppt-clr);
  }

  .dl-btn.pptx:hover {
    background: rgba(96, 165, 250, 0.18);
    border-color: var(--ppt-clr);
    transform: translateY(-2px);
    box-shadow: 0 6px 24px rgba(96, 165, 250, 0.2);
  }

  .btn-reset {
    background: transparent;
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 0.7rem;
    color: var(--muted);
    font-family: var(--body-font);
    font-size: 0.85rem;
    cursor: pointer;
    transition: all 0.2s ease;
    text-align: center;
    width: 100%;
  }

  .btn-reset:hover {
    border-color: var(--text);
    color: var(--text);
  }

  .error-msg {
    display: none;
    background: rgba(248, 113, 113, 0.08);
    border: 1px solid rgba(248, 113, 113, 0.25);
    border-radius: 10px;
    padding: 0.9rem 1.1rem;
    font-size: 0.85rem;
    color: var(--error);
    margin-top: 1.2rem;
    text-align: center;
  }

  .error-msg.show { display: block; }

  .footer {
    margin-top: 2.25rem;
    font-size: 0.78rem;
    color: var(--muted);
    text-align: center;
    line-height: 1.9;
  }

  .footer-copy {
    opacity: 0.55;
  }

  .footer-link {
    color: var(--accent);
    text-decoration: none;
    opacity: 0.8;
    transition: opacity 0.2s;
  }

  .footer-link:hover {
    opacity: 1;
    text-decoration: underline;
  }
' . "\n";
    echo '</style>' . "\n";
    echo '</head>' . "\n";
    echo '<body>' . "\n\n";

    echo '<div class="header">' . "\n";
    echo '  <h1>ZEX</h1>' . "\n";
    echo '  <p class="app-sub">Zip Export E<span class="x-mark">X</span>press</p>' . "\n";
    echo '  <p>' . $t['tag_line'] . '</p>' . "\n";
    echo '</div>' . "\n\n";

    echo '<div class="card">' . "\n";
    echo '  <div class="drop-zone" id="dropZone">
    <input type="file" id="fileInput" accept=".zip">
    <span class="drop-icon">' . "\xF0\x9F\x93\xA6" . '</span>
    <h3>Drop your ZIP here</h3>
    <p>Numbered high-quality PNGs &nbsp;&middot;&nbsp; click to browse &nbsp;&middot;&nbsp; max ' . MAX_MB . ' MB</p>
  </div>

  <div class="file-selected" id="fileSelected">
    <span class="icon">' . "\xF0\x9F\x97\x9C\xEF\xB8\x8F" . '</span>
    <div>
      <div class="name" id="fileName"></div>
      <div class="size" id="fileSize"></div>
    </div>
  </div>

  <div class="output-options" id="outputOptions">
    <span class="output-options-label">Output formats</span>
    <div class="output-options-row">
      <label class="output-option"><input type="checkbox" id="outPdf" name="out_pdf" value="1"> PDF</label>
      <label class="output-option"><input type="checkbox" id="outPptx" name="out_pptx" value="1"> PowerPoint (.pptx)</label>
    </div>
    <div class="output-options-sub">
      <label class="output-option output-option-supplement"><input type="checkbox" id="pdfWeb" name="pdf_web" value="1" disabled> Web PDF (smaller file)</label>
    </div>
  </div>

  <button class="btn-convert" id="convertBtn" disabled onclick="startConvert()">
    Convert Slides
  </button>

  <div class="progress-wrap" id="progressWrap">
    <div class="progress-bar-bg">
      <div class="progress-bar-fill" id="progressFill"></div>
    </div>
    <div class="progress-label" id="progressLabel">Uploading&hellip;</div>
  </div>

  <div class="error-msg" id="errorMsg"></div>

  <div class="results" id="results">
    <div class="results-header" id="resultsHeader"></div>
    <div class="download-btns" id="downloadBtns"></div>
    <button class="btn-reset" onclick="reset()">&larr; Convert another file</button>
  </div>
</div>

<div class="footer">
  <p>Files are stored temporarily and automatically cleaned up.</p>
  <p class="footer-copy">&copy; ' . date('Y') . ' ' . htmlspecialchars(COPYRIGHT_HOLDER, ENT_QUOTES, 'UTF-8') . '. All rights reserved.</p>
  <p><a href="https://github.com/josephbu/zex" class="footer-link" target="_blank" rel="noopener">Powered by ZEX</a></p>
</div>' . "\n";

    echo '
<script>
let selectedFile = null;

const dropZone     = document.getElementById("dropZone");
const fileInput    = document.getElementById("fileInput");
const fileSelected = document.getElementById("fileSelected");
const fileName     = document.getElementById("fileName");
const fileSize     = document.getElementById("fileSize");
const convertBtn   = document.getElementById("convertBtn");
const outPdfCb     = document.getElementById("outPdf");
const outPptxCb    = document.getElementById("outPptx");
const pdfWebCb     = document.getElementById("pdfWeb");
const progressWrap = document.getElementById("progressWrap");
const progressFill = document.getElementById("progressFill");
const progressLabel= document.getElementById("progressLabel");
const errorMsg     = document.getElementById("errorMsg");
const results      = document.getElementById("results");
const resultsHeader= document.getElementById("resultsHeader");
const downloadBtns = document.getElementById("downloadBtns");

function formatBytes(b) {
  if (b < 1024 * 1024) return (b / 1024).toFixed(1) + " KB";
  return (b / (1024 * 1024)).toFixed(1) + " MB";
}

function updateConvertEnabled() {
  const ok = selectedFile && (outPdfCb.checked || outPptxCb.checked);
  convertBtn.disabled = !ok;
}

function syncPdfWebControl() {
  if (!outPdfCb.checked) {
    pdfWebCb.checked = false;
    pdfWebCb.disabled = true;
  } else {
    pdfWebCb.disabled = false;
  }
}

function handleFile(file) {
  if (!file) return;
  if (!file.name.toLowerCase().endsWith(".zip")) {
    showError("Please select a .zip file.");
    return;
  }
  selectedFile = file;
  fileName.textContent = file.name;
  fileSize.textContent = formatBytes(file.size);
  fileSelected.classList.add("show");
  updateConvertEnabled();
  errorMsg.classList.remove("show");
}

fileInput.addEventListener("change", () => handleFile(fileInput.files[0]));
outPdfCb.addEventListener("change", () => { syncPdfWebControl(); updateConvertEnabled(); });
outPptxCb.addEventListener("change", updateConvertEnabled);
syncPdfWebControl();

dropZone.addEventListener("dragover", e => { e.preventDefault(); dropZone.classList.add("dragover"); });
dropZone.addEventListener("dragleave", () => dropZone.classList.remove("dragover"));
dropZone.addEventListener("drop", e => {
  e.preventDefault();
  dropZone.classList.remove("dragover");
  handleFile(e.dataTransfer.files[0]);
});

function showError(msg) {
  errorMsg.textContent = msg;
  errorMsg.classList.add("show");
  progressWrap.classList.remove("show");
  updateConvertEnabled();
}

function setProgress(pct, label) {
  progressFill.style.width = pct + "%";
  progressLabel.textContent = label;
}

function startConvert() {
  if (!selectedFile) return;
  if (!outPdfCb.checked && !outPptxCb.checked) return;
  convertBtn.disabled = true;
  progressWrap.classList.add("show");
  results.classList.remove("show");
  errorMsg.classList.remove("show");
  setProgress(20, "Uploading\u2026");

  const form = new FormData();
  form.append("zipfile", selectedFile);
  if (outPdfCb.checked) form.append("out_pdf", "1");
  if (outPdfCb.checked && pdfWebCb.checked) form.append("pdf_web", "1");
  if (outPptxCb.checked) form.append("out_pptx", "1");

  const xhr = new XMLHttpRequest();
  xhr.open("POST", "?action=convert", true);

  xhr.upload.onprogress = e => {
    if (e.lengthComputable) {
      const pct = Math.round((e.loaded / e.total) * 60);
      setProgress(pct, "Uploading\u2026 " + formatBytes(e.loaded) + " / " + formatBytes(e.total));
    }
  };

  let fakePct = 60;
  const fakeTimer = setInterval(() => {
    if (fakePct < 90) { fakePct += 3; setProgress(fakePct, "Generating files\u2026"); }
    else clearInterval(fakeTimer);
  }, 400);

  xhr.onload = function() {
    clearInterval(fakeTimer);
    if (xhr.status !== 200) { showError("Server error: " + xhr.status); return; }
    try {
      const data = JSON.parse(xhr.responseText);
      if (data.error) { showError(data.error); return; }
      setProgress(100, "Done!");
      setTimeout(() => { progressWrap.classList.remove("show"); showResults(data); }, 500);
    } catch(e) { showError("Unexpected server response."); }
  };

  xhr.onerror = () => { clearInterval(fakeTimer); showError("Network error. Please try again."); };
  xhr.send(form);
}

function showResults(data) {
  const count = data.slideCount || "?";
  resultsHeader.innerHTML = "<strong>\u2713 " + count + " slides converted</strong> \u2014 ready to download";
  downloadBtns.innerHTML = "";
  if (data.pdf) {
    const a = document.createElement("a");
    a.href = "?action=download&file=" + encodeURIComponent(data.pdf) + (data.pdfName ? "&name=" + encodeURIComponent(data.pdfName) : "");
    a.download = data.pdfName || "slides.pdf";
    a.className = "dl-btn pdf";
    a.innerHTML = "<span class=\"dl-icon\">\uD83D\uDCC4</span><span>Download PDF</span><span class=\"dl-type\">Adobe PDF \u00B7 all slides</span>";
    downloadBtns.appendChild(a);
  }
  if (data.pptx) {
    const a = document.createElement("a");
    a.href = "?action=download&file=" + encodeURIComponent(data.pptx) + (data.pptxName ? "&name=" + encodeURIComponent(data.pptxName) : "");
    a.download = data.pptxName || "slides.pptx";
    a.className = "dl-btn pptx";
    a.innerHTML = "<span class=\"dl-icon\">\uD83D\uDCCA</span><span>Download PPTX</span><span class=\"dl-type\">PowerPoint \u00B7 editable</span>";
    downloadBtns.appendChild(a);
  }
  results.classList.add("show");
  updateConvertEnabled();
}

function reset() {
  selectedFile = null;
  fileInput.value = "";
  fileSelected.classList.remove("show");
  results.classList.remove("show");
  progressWrap.classList.remove("show");
  errorMsg.classList.remove("show");
  outPdfCb.checked = false;
  outPptxCb.checked = false;
  pdfWebCb.checked = false;
  pdfWebCb.disabled = true;
  convertBtn.disabled = true;
  fileName.textContent = "";
  fileSize.textContent = "";
}
</script>
</body>
</html>';
}
?>
