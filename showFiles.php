<html lang="en" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
</head>
<body>

 <?php
$dir = "./upload/";


// اگر دکمه حذف زده شد
if ($_SERVER["REQUEST_METHOD"] == "POST" && isset($_POST["delete_all"])) {

    $files = glob($dir . "*");

    foreach ($files as $file) {
        if (is_file($file)) {
            unlink($file); // حذف فایل
        }
    }

    echo "<p style='color:red;'>همه فایل‌ها با موفقیت حذف شدند.</p>";
}
?>

<h2>لیست فایل‌ها:</h2>
<ul>
<?php
if (is_dir($dir)) {

    $files = glob($dir . "*");

    foreach ($files as $file) {
        if (is_file($file)) {
            $filename = basename($file);
            echo "<li><a href='$file' target='_blank'>$filename</a></li>";
        }
    }
}
?>
</ul>

<!-- دکمه حذف همه -->
<form method="post" onsubmit="return confirm('آیا مطمئن هستید که می‌خواهید همه فایل‌ها حذف شوند؟');">
    <button type="submit" name="delete_all" 
        style="background:red;color:white;padding:10px 20px;border:none;cursor:pointer;">
        حذف همه فایل‌ها
    </button>
</form>
</body>
</html>


