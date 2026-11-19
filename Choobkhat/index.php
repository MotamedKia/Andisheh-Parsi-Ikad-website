<?php
/*
 * سیستم حسابداری فروشگاه خدمات IT
 * زبان: فارسی | رابط کاربری: LTR | واحد پول: ریال ایران (IRR)
 *
 * نصب:
 * 1) این فایل را روی PHP 8.1+ قرار دهید.
 * 2) کنار این فایل یک فایل PNG با نام logo.png قرار دهید (ترجیحاً 446×140).
 * 3) PHP باید PDO_SQLite فعال داشته باشد.
 *
 * خروجی Excel به‌صورت Excel XML (.xls) تولید می‌شود و با Microsoft Excel / LibreOffice باز می‌شود.
 * خروجی PDF رسید با کتابخانه html2pdf.js در مرورگر ساخته می‌شود و برای چاپ A5 تنظیم شده است.
 */

declare(strict_types=1);
session_start();

$db = new PDO('sqlite:' . __DIR__ . '/accounting.sqlite');
$db->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$db->exec("PRAGMA foreign_keys = ON");

$db->exec("
CREATE TABLE IF NOT EXISTS customers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT UNIQUE,
    full_name TEXT NOT NULL,
    phone TEXT,
    mobile TEXT,
    email TEXT,
    national_id TEXT,
    address TEXT,
    postal_code TEXT,
    notes TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT UNIQUE NOT NULL,
    type TEXT NOT NULL CHECK(type IN ('service','product')),
    category TEXT,
    name TEXT NOT NULL,
    buy_price INTEGER DEFAULT 0,
    sell_price INTEGER DEFAULT 0,
    units INTEGER DEFAULT 0,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS receipts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    receipt_no TEXT UNIQUE NOT NULL,
    kind TEXT NOT NULL CHECK(kind IN ('sale','purchase')),
    customer_id INTEGER,
    date TEXT NOT NULL,
    pay_date TEXT,
    discount_type TEXT DEFAULT 'percent' CHECK(discount_type IN ('percent','amount')),
    discount_value REAL DEFAULT 0,
    extra_info TEXT,
    total INTEGER DEFAULT 0,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(customer_id) REFERENCES customers(id) ON DELETE SET NULL
);
CREATE TABLE IF NOT EXISTS receipt_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    receipt_id INTEGER NOT NULL,
    item_id INTEGER NOT NULL,
    item_name TEXT NOT NULL,
    item_code TEXT NOT NULL,
    item_type TEXT NOT NULL,
    qty INTEGER NOT NULL DEFAULT 1,
    unit_price INTEGER NOT NULL DEFAULT 0,
    discount_type TEXT DEFAULT 'percent' CHECK(discount_type IN ('percent','amount')),
    discount_value REAL DEFAULT 0,
    line_total INTEGER DEFAULT 0,
    FOREIGN KEY(receipt_id) REFERENCES receipts(id) ON DELETE CASCADE,
    FOREIGN KEY(item_id) REFERENCES items(id)
);
");

function h($v): string { return htmlspecialchars((string)$v, ENT_QUOTES, 'UTF-8'); }
function money($v): string { return number_format((int)$v); }
function post($key, $default='') { return $_POST[$key] ?? $default; }
function get($key, $default='') { return $_GET[$key] ?? $default; }

function calcDiscount(float $base, string $type, float $value): int {
    if ($type === 'percent') return (int)round($base * min(100, max(0, $value)) / 100);
    return (int)round(min($base, max(0, $value)));
}
function receiptTotal(array $rows): int {
    $sum = 0;
    foreach ($rows as $r) $sum += (int)$r['line_total'];
    return $sum;
}
function xls($filename, array $headers, array $rows): never {
    header('Content-Type: application/vnd.ms-excel; charset=UTF-8');
    header('Content-Disposition: attachment; filename="' . $filename . '"');
    echo "\xEF\xBB\xBF";
    echo '<?xml version="1.0" encoding="UTF-8"?>';
    echo '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">';
    echo '<Worksheet ss:Name="گزارش"><Table>';
    echo '<Row>'; foreach ($headers as $v) echo '<Cell><Data ss:Type="String">'.h($v).'</Data></Cell>'; echo '</Row>';
    foreach ($rows as $row) {
        echo '<Row>';
        foreach ($row as $v) {
            $type = is_numeric($v) ? 'Number' : 'String';
            echo '<Cell><Data ss:Type="'.$type.'">'.h($v).'</Data></Cell>';
        }
        echo '</Row>';
    }
    echo '</Table></Worksheet></Workbook>';
    exit;
}
function go(string $url): never { header('Location: '.$url); exit; }

$action = get('action');
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $action = post('action');
}

/* ---------- CRUD مشتری ---------- */
if ($action === 'save_customer') {
    $id = (int)post('id');
    $data = [post('code') ?: ('CUS-'.date('ymdHis')), post('full_name'), post('phone'), post('mobile'),
             post('email'), post('national_id'), post('address'), post('postal_code'), post('notes')];
    if ($id) {
        $st=$db->prepare("UPDATE customers SET code=?,full_name=?,phone=?,mobile=?,email=?,national_id=?,address=?,postal_code=?,notes=? WHERE id=?");
        $st->execute([...$data,$id]);
    } else {
        $st=$db->prepare("INSERT INTO customers(code,full_name,phone,mobile,email,national_id,address,postal_code,notes) VALUES(?,?,?,?,?,?,?,?,?)");
        $st->execute($data);
    }
    go('?page=customers');
}
if ($action === 'delete_customer') {
    $db->prepare("DELETE FROM customers WHERE id=?")->execute([(int)post('id')]);
    go('?page=customers');
}

/* ---------- CRUD آیتم ---------- */
if ($action === 'save_item') {
    $id=(int)post('id');
    $data=[post('code'),post('type','product'),post('category'),post('name'),(int)post('buy_price'),(int)post('sell_price'),(int)post('units')];
    if ($id) {
        $st=$db->prepare("UPDATE items SET code=?,type=?,category=?,name=?,buy_price=?,sell_price=?,units=? WHERE id=?");
        $st->execute([...$data,$id]);
    } else {
        $st=$db->prepare("INSERT INTO items(code,type,category,name,buy_price,sell_price,units) VALUES(?,?,?,?,?,?,?)");
        $st->execute($data);
    }
    go('?page='.($data[1]==='service'?'services':'products'));
}
if ($action === 'delete_item') {
    $db->prepare("DELETE FROM items WHERE id=?")->execute([(int)post('id')]);
    go('?page='.get('return','products'));
}

/* ---------- رسید ---------- */
if ($action === 'save_receipt') {
    $id=(int)post('id');
    $kind=post('kind','sale');
    $customer=(int)post('customer_id') ?: null;
    $date=post('date',date('Y-m-d'));
    $payDate=post('pay_date') ?: null;
    $dType=post('discount_type','percent');
    $dValue=(float)post('discount_value',0);
    $extra=post('extra_info');
    $items=json_decode(post('items_json','[]'), true) ?: [];

    try {
        $db->beginTransaction();

        if ($id) {
            $old=$db->prepare("SELECT * FROM receipt_items WHERE receipt_id=?");
            $old->execute([$id]);
            $oldItems=$old->fetchAll(PDO::FETCH_ASSOC);
            $oldReceipt=$db->prepare("SELECT kind FROM receipts WHERE id=?");
            $oldReceipt->execute([$id]);
            $oldKind=$oldReceipt->fetchColumn();

            // بازگرداندن موجودی رسید قبلی
            foreach ($oldItems as $oi) {
                if ($oi['item_type']==='product') {
                    $delta = ($oldKind==='sale') ? $oi['qty'] : -$oi['qty'];
                    $db->prepare("UPDATE items SET units=units+? WHERE id=?")->execute([$delta,$oi['item_id']]);
                }
            }
            $db->prepare("DELETE FROM receipt_items WHERE receipt_id=?")->execute([$id]);
        } else {
            $id=null;
        }

        $sub=0;
        $normalized=[];
        foreach ($items as $r) {
            $itemId=(int)($r['item_id']??0);
            $st=$db->prepare("SELECT * FROM items WHERE id=?");
            $st->execute([$itemId]); $it=$st->fetch(PDO::FETCH_ASSOC);
            if (!$it) continue;
            $qty=max(1,(int)($r['qty']??1));
            $unit=(int)($r['unit_price']??($kind==='sale'?$it['sell_price']:$it['buy_price']));
            $itDType=$r['discount_type']??'percent';
            $itDValue=(float)($r['discount_value']??0);
            $lineBase=$qty*$unit;
            $line=$lineBase-calcDiscount($lineBase,$itDType,$itDValue);
            $sub += $line;
            $normalized[]=[$it,$qty,$unit,$itDType,$itDValue,$line];
        }
        $total=$sub-calcDiscount($sub,$dType,$dValue);
        $receiptNo=post('receipt_no') ?: ('REC-'.date('Ymd-His'));

        if ($id) {
            $st=$db->prepare("UPDATE receipts SET receipt_no=?,kind=?,customer_id=?,date=?,pay_date=?,discount_type=?,discount_value=?,extra_info=?,total=? WHERE id=?");
            $st->execute([$receiptNo,$kind,$customer,$date,$payDate,$dType,$dValue,$extra,$total,$id]);
        } else {
            $st=$db->prepare("INSERT INTO receipts(receipt_no,kind,customer_id,date,pay_date,discount_type,discount_value,extra_info,total) VALUES(?,?,?,?,?,?,?,?,?)");
            $st->execute([$receiptNo,$kind,$customer,$date,$payDate,$dType,$dValue,$extra,$total]);
            $id=(int)$db->lastInsertId();
        }

        foreach ($normalized as [$it,$qty,$unit,$itDType,$itDValue,$line]) {
            $st=$db->prepare("INSERT INTO receipt_items(receipt_id,item_id,item_name,item_code,item_type,qty,unit_price,discount_type,discount_value,line_total) VALUES(?,?,?,?,?,?,?,?,?,?)");
            $st->execute([$id,$it['id'],$it['name'],$it['code'],$it['type'],$qty,$unit,$itDType,$itDValue,$line]);
            if ($it['type']==='product') {
                // فروش: کاهش موجودی | خرید: افزایش موجودی
                $delta=($kind==='sale') ? -$qty : $qty;
                $db->prepare("UPDATE items SET units=units+? WHERE id=?")->execute([$delta,$it['id']]);
            }
        }
        $db->commit();
        go('?page=receipts&saved=1');
    } catch(Throwable $e) {
        if ($db->inTransaction()) $db->rollBack();
        $_SESSION['error']=$e->getMessage();
        go('?page=receipt_add');
    }
}
if ($action === 'delete_receipt') {
    $id=(int)post('id');
    try {
        $db->beginTransaction();
        $r=$db->prepare("SELECT kind FROM receipts WHERE id=?"); $r->execute([$id]); $kind=$r->fetchColumn();
        $st=$db->prepare("SELECT item_id,item_type,qty FROM receipt_items WHERE receipt_id=?"); $st->execute([$id]);
        foreach($st as $x) if($x['item_type']==='product') {
            $delta=($kind==='sale')?$x['qty']:-$x['qty'];
            $db->prepare("UPDATE items SET units=units+? WHERE id=?")->execute([$delta,$x['item_id']]);
        }
        $db->prepare("DELETE FROM receipts WHERE id=?")->execute([$id]);
        $db->commit();
    } catch(Throwable $e){ if($db->inTransaction())$db->rollBack(); }
    go('?page=receipts');
}

/* ---------- Excel exports ---------- */
if ($action === 'excel_customers') {
    $rows=$db->query("SELECT code,full_name,phone,mobile,email,national_id,address,postal_code,notes FROM customers ORDER BY id DESC")->fetchAll(PDO::FETCH_NUM);
    xls('customers.xls',['کد','نام کامل','تلفن','موبایل','ایمیل','کد ملی','آدرس','کد پستی','توضیحات'],$rows);
}
if ($action === 'excel_items') {
    $type=get('type','product');
    $st=$db->prepare("SELECT code,category,name,buy_price,sell_price,units,type FROM items WHERE type=? ORDER BY id DESC"); $st->execute([$type]);
    xls(($type==='service'?'services':'products').'.xls',['شناسه','دسته‌بندی','نام','قیمت خرید','قیمت فروش','موجودی/واحد','نوع'],$st->fetchAll(PDO::FETCH_NUM));
}
if ($action === 'excel_receipts') {
    $sql="SELECT r.receipt_no,r.kind,r.date,r.pay_date,COALESCE(c.full_name,'-'),r.discount_type,r.discount_value,r.total,r.extra_info
          FROM receipts r LEFT JOIN customers c ON c.id=r.customer_id WHERE 1=1";
    $params=[];
    foreach(['receipt_no','kind','date','pay_date'] as $f) if(get($f)!==''){ $sql.=" AND r.$f LIKE ?"; $params[]='%'.get($f).'%'; }
    if(get('customer')!==''){ $sql.=" AND c.full_name LIKE ?"; $params[]='%'.get('customer').'%'; }
    $sql.=" ORDER BY r.id DESC";
    $st=$db->prepare($sql);$st->execute($params);
    xls('receipts.xls',['شماره رسید','نوع','تاریخ','تاریخ پرداخت','مشتری','نوع تخفیف','مقدار تخفیف','جمع نهایی','توضیحات'],$st->fetchAll(PDO::FETCH_NUM));
}


/* ---------- چاپ / PDF رسید ---------- */
if ($action === 'receipt_print') {
    $id=(int)get('id');
    $st=$db->prepare("SELECT r.*,COALESCE(c.full_name,'بدون مشتری') customer_name,COALESCE(c.mobile,'') customer_mobile FROM receipts r LEFT JOIN customers c ON c.id=r.customer_id WHERE r.id=?");
    $st->execute([$id]); $r=$st->fetch(PDO::FETCH_ASSOC);
    if(!$r){http_response_code(404);exit('رسید یافت نشد');}
    $st=$db->prepare("SELECT * FROM receipt_items WHERE receipt_id=?");$st->execute([$id]);$ri=$st->fetchAll(PDO::FETCH_ASSOC);
    ?>
    <!doctype html><html lang="fa" dir="rtl"><head><meta charset="utf-8"><title>رسید <?=h($r['receipt_no'])?></title>
    <link href="https://fonts.googleapis.com/css2?family=Vazirmatn:wght@400;600;700&display=swap" rel="stylesheet">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
    <style>@page{size:A5;margin:0}body{margin:0;font-family:Vazirmatn,Tahoma,sans-serif;background:#eee}.toolbar{text-align:center;padding:15px;direction:rtl}.btn{padding:10px 16px;border:0;border-radius:7px;background:#1591D4;color:#fff;cursor:pointer}.btn.o{background:#E3783D}
    #receiptPrint{direction:rtl;background:#fff;width:148mm;min-height:210mm;margin:15px auto;padding:8mm;box-sizing:border-box}#receiptPrint img{width:90mm;max-height:28mm;object-fit:contain;display:block;margin:0 auto 5mm}.head{border-bottom:2px solid #1591D4;padding-bottom:5mm}.row{display:flex;justify-content:space-between;margin:3mm 0}table{width:100%;border-collapse:collapse;font-size:11px}th,td{padding:5px;border-bottom:1px solid #ddd;text-align:right}th{background:#f0f7fb}.total{font-size:16px;font-weight:bold;color:#1591D4;margin-top:7mm}.extra{margin-top:7mm;font-size:11px}@media print{body{background:#fff}.toolbar{display:none}#receiptPrint{margin:0}}</style></head><body>
    <div class="toolbar"><button class="btn" onclick="makePDF()">دریافت PDF A5</button> <button class="btn o" onclick="window.print()">چاپ / ذخیره PDF</button></div>
    <section id="receiptPrint">
      <img src="logo.png" alt="لوگو">
      <div class="head"><div class="row"><b>رسید: <?=h($r['receipt_no'])?></b><b><?=($r['kind']==='sale'?'فروش':'خرید')?></b></div>
      <div class="row"><span>تاریخ: <?=h($r['date'])?></span><span>تاریخ پرداخت: <?=h($r['pay_date'] ?: '-')?></span></div>
      <div class="row"><span>مشتری: <?=h($r['customer_name'])?></span><span><?=h($r['customer_mobile'])?></span></div></div>
      <table><thead><tr><th>شرح</th><th>تعداد</th><th>قیمت واحد</th><th>تخفیف</th><th>مبلغ</th></tr></thead><tbody>
      <?php foreach($ri as $x): ?><tr><td><?=h($x['item_name'])?><br><small><?=h($x['item_code'])?></small></td><td><?=h($x['qty'])?></td><td><?=money($x['unit_price'])?></td><td><?=h($x['discount_value']).($x['discount_type']==='percent'?'٪':' ریال')?></td><td><?=money($x['line_total'])?></td></tr><?php endforeach;?>
      </tbody></table>
      <div class="row"><span>تخفیف کل:</span><span><?=h($r['discount_value']).($r['discount_type']==='percent'?'٪':' ریال')?></span></div>
      <div class="total">مبلغ نهایی: <?=money($r['total'])?> ریال</div>
      <?php if($r['extra_info']): ?><div class="extra"><b>توضیحات:</b> <?=nl2br(h($r['extra_info']))?></div><?php endif;?>
    </section>
    <script>function makePDF(){html2pdf().set({margin:0,filename:'<?=h($r['receipt_no'])?>.pdf',image:{type:'jpeg',quality:.98},html2canvas:{scale:2,useCORS:true},jsPDF:{unit:'mm',format:'a5',orientation:'portrait'}}).from(document.getElementById('receiptPrint')).save()}</script>
    </body></html><?php exit;
}

/* ---------- داده‌های نمایش ---------- */
$page=get('page','dashboard');
$customers=$db->query("SELECT * FROM customers ORDER BY id DESC")->fetchAll(PDO::FETCH_ASSOC);

function itemList(PDO $db,string $type): array {
    $st=$db->prepare("SELECT * FROM items WHERE type=? ORDER BY id DESC");$st->execute([$type]);return $st->fetchAll(PDO::FETCH_ASSOC);
}
$services=itemList($db,'service');
$products=itemList($db,'product');

function filteredReceipts(PDO $db): array {
    $sql="SELECT r.*,COALESCE(c.full_name,'-') customer_name FROM receipts r LEFT JOIN customers c ON c.id=r.customer_id WHERE 1=1";
    $p=[];
    foreach(['receipt_no','kind','date','pay_date'] as $f) if(get($f)!==''){ $sql.=" AND r.$f LIKE ?";$p[]='%'.get($f).'%';}
    if(get('customer')!==''){ $sql.=" AND c.full_name LIKE ?";$p[]='%'.get('customer').'%';}
    $sql.=" ORDER BY r.id DESC";
    $st=$db->prepare($sql);$st->execute($p);return $st->fetchAll(PDO::FETCH_ASSOC);
}
$receipts=filteredReceipts($db);

$editCustomer=null;
if($page==='customers' && get('edit')){$st=$db->prepare("SELECT * FROM customers WHERE id=?");$st->execute([(int)get('edit')]);$editCustomer=$st->fetch(PDO::FETCH_ASSOC);}
$editItem=null;
if(in_array($page,['services','products']) && get('edit')){$st=$db->prepare("SELECT * FROM items WHERE id=?");$st->execute([(int)get('edit')]);$editItem=$st->fetch(PDO::FETCH_ASSOC);}

$editReceipt=null;$editReceiptItems=[];
if($page==='receipt_add' && get('edit')){
    $st=$db->prepare("SELECT * FROM receipts WHERE id=?");$st->execute([(int)get('edit')]);$editReceipt=$st->fetch(PDO::FETCH_ASSOC);
    if($editReceipt){$st=$db->prepare("SELECT * FROM receipt_items WHERE receipt_id=?");$st->execute([$editReceipt['id']]);$editReceiptItems=$st->fetchAll(PDO::FETCH_ASSOC);}
}
?>
<!doctype html>
<html lang="fa" dir="ltr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>حسابداری فروشگاه IT</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Vazirmatn:wght@400;500;600;700&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
<style>
*{box-sizing:border-box}body{margin:0;background:#fff;color:#17212b;font-family:Vazirmatn,Tahoma,sans-serif}
button,input,select,textarea{font-family:inherit}a{text-decoration:none;color:inherit}
.layout{min-height:100vh;display:flex;flex-direction:row-reverse}
.sidebar{width:20%;min-width:230px;background:#fff;border-left:1px solid #e4e8ec;padding:18px;position:sticky;top:0;height:100vh;overflow:auto}
.sidebar img{width:100%;max-width:446px;height:auto;display:block;margin:0 auto 20px}
.brand{font-weight:800;color:#1591D4;text-align:center;margin-bottom:18px}
.nav a{display:block;padding:12px 14px;border-radius:10px;margin:5px 0;color:#34424d}
.nav a:hover,.nav a.active{background:#1591D4;color:#fff}.nav .sep{height:1px;background:#e9edf0;margin:14px 0}
.content{width:80%;padding:28px;direction:rtl;overflow:hidden}
.menu-btn{display:none;position:fixed;right:14px;top:12px;z-index:30;background:#1591D4;color:white;border:0;border-radius:10px;padding:10px 13px;font-size:20px}
h1{margin-top:0;font-size:26px;color:#1591D4}.toolbar{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:16px;align-items:center}
.btn{border:0;border-radius:8px;padding:9px 14px;cursor:pointer;background:#1591D4;color:#fff;font-weight:600}.btn.orange{background:#E3783D}.btn.gray{background:#687782}.btn.danger{background:#c83c3c}.btn.light{background:#edf6fb;color:#1591D4}
.card{background:#fff;border:1px solid #e5eaee;border-radius:14px;padding:18px;margin-bottom:18px;box-shadow:0 3px 15px rgba(20,50,70,.05)}
.grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}.grid.two{grid-template-columns:repeat(2,1fr)}label{display:block;font-size:13px;font-weight:600;margin-bottom:5px}
input,select,textarea{width:100%;padding:9px 10px;border:1px solid #ccd6dd;border-radius:8px;background:#fff;direction:rtl}textarea{min-height:80px}
table{width:100%;border-collapse:collapse;font-size:13px}th,td{padding:9px 8px;border-bottom:1px solid #edf0f2;text-align:right;vertical-align:middle}th{background:#f5f9fb;color:#42515c;position:sticky;top:0}.table-wrap{overflow:auto;max-height:62vh}
.badge{padding:4px 8px;border-radius:999px;background:#eaf6fc;color:#0877ae}.badge.orange{background:#fff1e9;color:#bd5726}
.stats{display:grid;grid-template-columns:repeat(3,1fr);gap:15px}.stat{padding:20px;border-radius:14px;background:#f5faff;border-right:5px solid #1591D4}.stat b{font-size:22px}
.receipt-items td input,.receipt-items td select{padding:6px}.remove{border:0;background:#fee;color:#b22;border-radius:7px;padding:7px;cursor:pointer}
.search-grid{display:grid;grid-template-columns:repeat(5,1fr);gap:8px}.muted{color:#73808a;font-size:12px}.total{font-size:20px;font-weight:800;color:#1591D4}
#receiptPrint{direction:rtl;background:white;width:148mm;min-height:210mm;margin:auto;padding:8mm;font-family:Vazirmatn,Tahoma,sans-serif}
#receiptPrint img{width:90mm;max-height:28mm;object-fit:contain;display:block;margin:0 auto 5mm}
.print-only{display:none}
@media(max-width:900px){.sidebar{width:260px;position:fixed;right:-270px;z-index:20;transition:.25s}.sidebar.open{right:0}.content{width:100%;padding:60px 15px 20px}.menu-btn{display:block}.grid,.stats{grid-template-columns:1fr}.search-grid{grid-template-columns:repeat(2,1fr)}}
@media print{body{background:#fff}.sidebar,.content>h1,.content>.toolbar,.no-print{display:none!important}.content{width:100%;padding:0}.print-only{display:block}}
</style>
</head>
<body>
<button class="menu-btn" onclick="document.querySelector('.sidebar').classList.toggle('open')">☰</button>
<div class="layout">
<aside class="sidebar">
<img src="logo.png" alt="لوگو" onerror="this.style.display='none'">
<div class="brand">حسابداری فروشگاه IT</div>
<nav class="nav">
<a class="<?=($page==='dashboard'?'active':'')?>" href="?page=dashboard">داشبورد</a>
<a class="<?=($page==='receipts'?'active':'')?>" href="?page=receipts">لیست رسیدها</a>
<a class="<?=($page==='receipt_add'?'active':'')?>" href="?page=receipt_add">افزودن رسید</a>
<div class="sep"></div>
<a class="<?=($page==='services'?'active':'')?>" href="?page=services">لیست خدمات</a>
<a class="<?=($page==='products'?'active':'')?>" href="?page=products">لیست محصولات</a>
<div class="sep"></div>
<a class="<?=($page==='customers'?'active':'')?>" href="?page=customers">لیست مشتریان</a>
</nav>
</aside>

<main class="content">
<?php if(!empty($_SESSION['error'])): ?><div class="card" style="color:#b22"><?=h($_SESSION['error']);unset($_SESSION['error']);?></div><?php endif; ?>

<?php if($page==='dashboard'): ?>
<h1>داشبورد</h1>
<div class="stats">
<div class="stat"><div>تعداد مشتریان</div><b><?=count($customers)?></b></div>
<div class="stat"><div>تعداد خدمات</div><b><?=count($services)?></b></div>
<div class="stat"><div>تعداد محصولات</div><b><?=count($products)?></b></div>
</div>
<div class="card"><h3>راهنمای سریع</h3><p>برای ثبت فروش یا خرید، ابتدا مشتری و سپس خدمات/محصولات را تعریف کنید. در رسید، نوع «فروش» موجودی محصول را کم و نوع «خرید» موجودی را زیاد می‌کند.</p></div>

<?php elseif($page==='customers'): ?>
<h1>مشتریان</h1>
<div class="toolbar"><a class="btn" href="?page=customers">مشتری جدید</a><a class="btn orange" href="?action=excel_customers">خروجی Excel</a></div>
<div class="card">
<h3><?= $editCustomer?'ویرایش مشتری':'افزودن مشتری' ?></h3>
<form method="post"><input type="hidden" name="action" value="save_customer"><input type="hidden" name="id" value="<?=h($editCustomer['id']??'')?>">
<div class="grid">
<div><label>کد</label><input name="code" value="<?=h($editCustomer['code']??'')?>"></div>
<div><label>نام کامل *</label><input required name="full_name" value="<?=h($editCustomer['full_name']??'')?>"></div>
<div><label>تلفن</label><input name="phone" value="<?=h($editCustomer['phone']??'')?>"></div>
<div><label>موبایل</label><input name="mobile" value="<?=h($editCustomer['mobile']??'')?>"></div>
<div><label>ایمیل</label><input type="email" name="email" value="<?=h($editCustomer['email']??'')?>"></div>
<div><label>کد ملی</label><input name="national_id" value="<?=h($editCustomer['national_id']??'')?>"></div>
<div><label>کد پستی</label><input name="postal_code" value="<?=h($editCustomer['postal_code']??'')?>"></div>
<div style="grid-column:span 2"><label>آدرس</label><input name="address" value="<?=h($editCustomer['address']??'')?>"></div>
<div style="grid-column:span 3"><label>توضیحات</label><textarea name="notes"><?=h($editCustomer['notes']??'')?></textarea></div>
</div><br><button class="btn">ذخیره</button></form>
</div>
<div class="card"><h3>جستجو</h3><div class="search-grid"><input class="live-filter" placeholder="کد" data-col="0"><input class="live-filter" placeholder="نام" data-col="1"><input class="live-filter" placeholder="تلفن" data-col="2"><input class="live-filter" placeholder="موبایل" data-col="3"><input class="live-filter" placeholder="ایمیل" data-col="4"></div></div>
<div class="card table-wrap"><table class="filter-table"><thead><tr><th>کد</th><th>نام کامل</th><th>تلفن</th><th>موبایل</th><th>ایمیل</th><th>کد ملی</th><th>عملیات</th></tr></thead><tbody>
<?php foreach($customers as $c): ?><tr><td><?=h($c['code'])?></td><td><?=h($c['full_name'])?></td><td><?=h($c['phone'])?></td><td><?=h($c['mobile'])?></td><td><?=h($c['email'])?></td><td><?=h($c['national_id'])?></td><td><a class="btn light" href="?page=customers&edit=<?=$c['id']?>">ویرایش</a> <form style="display:inline" method="post" onsubmit="return confirm('حذف شود؟')"><input type="hidden" name="action" value="delete_customer"><input type="hidden" name="id" value="<?=$c['id']?>"><button class="btn danger">حذف</button></form></td></tr><?php endforeach;?>
</tbody></table></div>

<?php elseif($page==='services' || $page==='products'):
$type=$page==='services'?'service':'product';$list=$type==='service'?$services:$products;$title=$type==='service'?'خدمات':'محصولات'; ?>
<h1>لیست <?=$title?></h1>
<div class="toolbar"><a class="btn" href="?page=<?=$page?>">مورد جدید</a><a class="btn orange" href="?action=excel_items&type=<?=$type?>">خروجی Excel</a></div>
<div class="card"><h3><?= $editItem?'ویرایش':'افزودن'?> <?=$title?></h3>
<form method="post"><input type="hidden" name="action" value="save_item"><input type="hidden" name="id" value="<?=h($editItem['id']??'')?>"><input type="hidden" name="type" value="<?=$type?>">
<div class="grid">
<div><label>شناسه یکتا *</label><input required name="code" value="<?=h($editItem['code']??'')?>"></div>
<div><label>دسته‌بندی</label><input name="category" value="<?=h($editItem['category']??'')?>"></div>
<div><label>نام *</label><input required name="name" value="<?=h($editItem['name']??'')?>"></div>
<div><label>قیمت خرید (ریال)</label><input type="number" min="0" name="buy_price" value="<?=h($editItem['buy_price']??0)?>"></div>
<div><label>قیمت فروش (ریال)</label><input type="number" min="0" name="sell_price" value="<?=h($editItem['sell_price']??0)?>"></div>
<div><label>واحد / موجودی</label><input type="number" min="0" name="units" value="<?=h($editItem['units']??0)?>"></div>
</div><br><button class="btn">ذخیره</button></form></div>
<div class="card"><h3>جستجو</h3><div class="search-grid"><input class="live-filter" placeholder="شناسه" data-col="0"><input class="live-filter" placeholder="دسته‌بندی" data-col="1"><input class="live-filter" placeholder="نام" data-col="2"><input class="live-filter" placeholder="قیمت خرید" data-col="3"><input class="live-filter" placeholder="قیمت فروش" data-col="4"><input class="live-filter" placeholder="واحد/موجودی" data-col="5"></div></div>
<div class="card table-wrap"><table class="filter-table"><thead><tr><th>شناسه</th><th>دسته‌بندی</th><th>نام</th><th>خرید</th><th>فروش</th><th>واحد/موجودی</th><th>عملیات</th></tr></thead><tbody>
<?php foreach($list as $it): ?><tr><td><?=h($it['code'])?></td><td><?=h($it['category'])?></td><td><?=h($it['name'])?></td><td><?=money($it['buy_price'])?></td><td><?=money($it['sell_price'])?></td><td><?=h($it['units'])?></td><td><a class="btn light" href="?page=<?=$page?>&edit=<?=$it['id']?>">ویرایش</a> <form style="display:inline" method="post" onsubmit="return confirm('حذف شود؟')"><input type="hidden" name="action" value="delete_item"><input type="hidden" name="id" value="<?=$it['id']?>"><input type="hidden" name="return" value="<?=$page?>"><button class="btn danger">حذف</button></form></td></tr><?php endforeach;?>
</tbody></table></div>

<?php elseif($page==='receipts'): ?>
<h1>لیست رسیدها</h1>
<div class="toolbar"><a class="btn" href="?page=receipt_add">رسید جدید</a><a class="btn orange" href="?action=excel_receipts&<?=http_build_query($_GET)?>">خروجی Excel</a></div>
<div class="card"><h3>جستجوی چندمعیاره</h3>
<form><input type="hidden" name="page" value="receipts"><div class="search-grid">
<input name="receipt_no" value="<?=h(get('receipt_no'))?>" placeholder="شماره رسید">
<input name="kind" value="<?=h(get('kind'))?>" placeholder="نوع: sale / purchase">
<input name="date" value="<?=h(get('date'))?>" placeholder="تاریخ">
<input name="pay_date" value="<?=h(get('pay_date'))?>" placeholder="تاریخ پرداخت">
<input name="customer" value="<?=h(get('customer'))?>" placeholder="مشتری">
</div><br><button class="btn">جستجو</button></form></div>
<div class="card table-wrap"><table><thead><tr><th>شماره</th><th>نوع</th><th>مشتری</th><th>تاریخ</th><th>تاریخ پرداخت</th><th>جمع نهایی</th><th>عملیات</th></tr></thead><tbody>
<?php foreach($receipts as $r): ?><tr><td><?=h($r['receipt_no'])?></td><td><span class="badge <?=$r['kind']==='purchase'?'orange':''?>"><?=$r['kind']==='sale'?'فروش':'خرید'?></span></td><td><?=h($r['customer_name'])?></td><td><?=h($r['date'])?></td><td><?=h($r['pay_date'])?></td><td><?=money($r['total'])?> ریال</td><td><a class="btn light" href="?page=receipt_add&edit=<?=$r['id']?>">ویرایش</a> <a class="btn orange" target="_blank" href="?action=receipt_print&id=<?=$r['id']?>">PDF / چاپ A5</a> <form style="display:inline" method="post" onsubmit="return confirm('رسید حذف و موجودی اصلاح شود؟')"><input type="hidden" name="action" value="delete_receipt"><input type="hidden" name="id" value="<?=$r['id']?>"><button class="btn danger">حذف</button></form></td></tr><?php endforeach;?>
</tbody></table></div>

<?php elseif($page==='receipt_add'): ?>
<h1><?= $editReceipt?'ویرایش رسید':'افزودن رسید'?></h1>
<div class="card">
<form id="receiptForm" method="post">
<input type="hidden" name="action" value="save_receipt"><input type="hidden" name="id" value="<?=h($editReceipt['id']??'')?>"><input type="hidden" name="items_json" id="items_json">
<div class="grid">
<div><label>شماره رسید</label><input name="receipt_no" value="<?=h($editReceipt['receipt_no']??'')?>"></div>
<div><label>نوع رسید</label><select name="kind" id="kind"><option value="sale" <?=($editReceipt['kind']??'sale')==='sale'?'selected':''?>>فروش</option><option value="purchase" <?=($editReceipt['kind']??'')==='purchase'?'selected':''?>>خرید</option></select></div>
<div><label>مشتری</label><select name="customer_id"><option value="">بدون مشتری</option><?php foreach($customers as $c):?><option value="<?=$c['id']?>" <?=((int)($editReceipt['customer_id']??0)==$c['id']?'selected':'')?>><?=h($c['full_name'])?> — <?=h($c['mobile'])?></option><?php endforeach;?></select></div>
<div><label>تاریخ</label><input type="date" name="date" value="<?=h($editReceipt['date']??date('Y-m-d'))?>"></div>
<div><label>تاریخ پرداخت</label><input type="date" name="pay_date" value="<?=h($editReceipt['pay_date']??'')?>"></div>
<div><label>تخفیف کل: نوع</label><select name="discount_type"><option value="percent" <?=($editReceipt['discount_type']??'percent')==='percent'?'selected':''?>>درصدی</option><option value="amount" <?=($editReceipt['discount_type']??'')==='amount'?'selected':''?>>مبلغی (ریال)</option></select></div>
<div><label>مقدار تخفیف کل</label><input type="number" min="0" name="discount_value" value="<?=h($editReceipt['discount_value']??0)?>"></div>
<div style="grid-column:span 2"><label>توضیحات اضافی (اختیاری)</label><textarea name="extra_info"><?=h($editReceipt['extra_info']??'')?></textarea></div>
</div>

<hr><h3>اقلام رسید</h3>
<div class="card" style="background:#f8fbfd">
<div class="grid two"><div><label>جستجوی کالا/خدمت با شناسه یا نام</label><input id="itemSearch" placeholder="مثلاً: 200 گرم یا PR-102"></div><div><label>نتیجه</label><select id="itemPicker"><option value="">ابتدا جستجو کنید</option></select></div></div><br><button type="button" class="btn" onclick="addPickedItem()">افزودن به رسید</button>
</div>
<div class="table-wrap"><table class="receipt-items"><thead><tr><th>شناسه</th><th>شرح</th><th>نوع</th><th>تعداد</th><th>قیمت واحد (ریال)</th><th>تخفیف آیتم</th><th>مبلغ</th><th></th></tr></thead><tbody id="receiptRows"></tbody></table></div>
<div style="text-align:left;margin-top:15px"><span class="total">جمع: <span id="grandTotal">0</span> ریال</span></div><br>
<button class="btn orange" type="submit">ذخیره رسید</button>
</form></div>

<?php endif; ?>
</main></div>

<script>
const allItems = <?=json_encode(array_merge($services,$products),JSON_UNESCAPED_UNICODE)?>;
let receiptItems = <?=json_encode(array_map(function($r){return [
'item_id'=>(int)$r['item_id'],'item_code'=>$r['item_code'],'item_name'=>$r['item_name'],'item_type'=>$r['item_type'],
'qty'=>(int)$r['qty'],'unit_price'=>(int)$r['unit_price'],'discount_type'=>$r['discount_type'],'discount_value'=>(float)$r['discount_value']
];},$editReceiptItems),JSON_UNESCAPED_UNICODE)?>;

document.querySelectorAll('.live-filter').forEach(inp=>{
 inp.addEventListener('input',()=>{
  const table=inp.closest('.card').nextElementSibling?.querySelector('table') || document.querySelector('.filter-table');
  if(!table)return;
  const filters=[...inp.closest('.card').querySelectorAll('.live-filter')];
  [...table.querySelectorAll('tbody tr')].forEach(tr=>{
   const ok=filters.every(f=>!f.value || (tr.children[+f.dataset.col]?.innerText||'').toLowerCase().includes(f.value.toLowerCase()));
   tr.style.display=ok?'':'none';
  });
 });
});
function searchItems(){
 const q=(document.getElementById('itemSearch')?.value||'').trim().toLowerCase();
 const sel=document.getElementById('itemPicker'); if(!sel)return;
 sel.innerHTML='<option value="">انتخاب کنید</option>';
 allItems.filter(x=>(x.code+' '+x.name+' '+(x.category||'')).toLowerCase().includes(q)).slice(0,100).forEach(x=>{
  const o=document.createElement('option');o.value=x.id;
  o.textContent=x.code+' — '+x.name+' — '+(x.type==='service'?'خدمت':'محصول');
  sel.appendChild(o);
 });
}
document.getElementById('itemSearch')?.addEventListener('input',searchItems);
function addPickedItem(){
 const id=+(document.getElementById('itemPicker')?.value||0),it=allItems.find(x=>+x.id===id);if(!it)return;
 const kind=document.getElementById('kind').value;
 const existing=receiptItems.find(x=>+x.item_id===id);
 if(existing) existing.qty++; else receiptItems.push({item_id:+it.id,item_code:it.code,item_name:it.name,item_type:it.type,qty:1,unit_price:kind==='sale'?+it.sell_price:+it.buy_price,discount_type:'percent',discount_value:0});
 renderReceipt();
}
function renderReceipt(){
 const body=document.getElementById('receiptRows');if(!body)return;body.innerHTML='';
 let total=0;
 receiptItems.forEach((r,i)=>{
 const base = r.qty * r.unit_price;

const disc = r.discount_type === 'percent'
    ? base * Math.min(100, r.discount_value) / 100
    : Math.min(base, r.qty * r.discount_value);

const line = Math.max(0, base - disc);
  const tr=document.createElement('tr');
  tr.innerHTML=`<td>${esc(r.item_code)}</td><td>${esc(r.item_name)}</td><td>${r.item_type==='service'?'خدمت':'محصول'}</td>
  <td><input type="number" min="1" value="${r.qty}" onchange="receiptItems[${i}].qty=Math.max(1,+this.value);renderReceipt()"></td>
  <td><input type="number" min="0" value="${r.unit_price}" onchange="receiptItems[${i}].unit_price=Math.max(0,+this.value);renderReceipt()"></td>
  <td><div style="display:flex;gap:4px"><select onchange="receiptItems[${i}].discount_type=this.value;renderReceipt()"><option value="percent" ${r.discount_type==='percent'?'selected':''}>%</option><option value="amount" ${r.discount_type==='amount'?'selected':''}>ریال</option></select><input type="number" min="0" value="${r.discount_value}" onchange="receiptItems[${i}].discount_value=Math.max(0,+this.value);renderReceipt()"></div></td>
  <td>${Math.round(line).toLocaleString()} ریال</td><td><button type="button" class="remove" onclick="receiptItems.splice(${i},1);renderReceipt()">حذف</button></td>`;
  body.appendChild(tr);
 });
 document.getElementById('grandTotal').textContent=Math.round(total).toLocaleString();
 document.getElementById('items_json').value=JSON.stringify(receiptItems);
}
function esc(s){return String(s??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[m]))}
document.getElementById('kind')?.addEventListener('change',()=>{
 const kind=document.getElementById('kind').value;
 receiptItems.forEach(r=>{const it=allItems.find(x=>+x.id===+r.item_id);if(it)r.unit_price=kind==='sale'?+it.sell_price:+it.buy_price});
 renderReceipt();
});
renderReceipt();
</script>
</body></html>
