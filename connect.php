<?php
require_once('vendor/autoload.php');
require_once('connect.php');
?>
<!DOCTYPE>
<html>
<body>
<div>
    <form action="import.php" method="POST" enctype="multipart/form-data">
        <label for="myfile">Chọn file tải lên ở đây</label>
        <input type="file" id="myfile"  name="file"  ><br><br>
    </form>
</div>
<div class="export">
    <form action="exportExcel.php" method="POST" enctype="multipart/form-data">
        <input type="submit" value="exportExcel" name="export">
    </form>
</div>
</body>
</html>
<style>
    .export{
        margin-top: 50px;
    }
</style>