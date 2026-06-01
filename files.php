<?php
if (isset($_FILES['file'])) {

    $name = time() ."_". $_FILES['file']['name'];
    $tmp = $_FILES['file']['tmp_name'];

    if (move_uploaded_file($tmp, "upload/$name")) {
        echo "Upload Successfully";
    } else {
        echo "Error : " . $_FILES['file']['error'];
    }
}
?>

<DOCTYPE html>
    <html lang="fa" dir="rtl">

    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Send files</title>

        <style>
            body {
                text-align: right !important;
            }
        </style>
    </head>

    <body>
        <div>
            <form method="post" enctype="multipart/form-data">

                <input type="file" name="file">

                <button type="submit" name="send">Send</button>
            </form>
        </div>
    </body>

    </html>
</DOCTYPE>