<?php

$rows = [
    ['url', 'meta_title', 'meta_description'],
    ['/page-1', 'Best flight deals today', 'Find cheap flights and save big on your next trip with our exclusive deals.'],
    ['/page-2', '', 'Get compensation for delayed or cancelled flights quickly and easily.'],
    ['/page-3', 'Book cheap hotels now', ''],
    ['/page-4', 'Affordable car rentals', 'Rent a car at the best price. Compare hundreds of deals in seconds.'],
];

header('Content-Type: text/csv; charset=UTF-8');
header('Content-Disposition: attachment; filename="meta-template.csv"');

// UTF-8 BOM so Excel opens it correctly
echo "\xEF\xBB\xBF";

$out = fopen('php://output', 'w');
foreach ($rows as $row) {
    fputcsv($out, $row);
}
fclose($out);
