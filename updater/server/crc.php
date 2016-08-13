<?
//define the path as relative
$path = "files";
//using the opendir function
$dir_handle = @opendir($path) or die("Unable to open $path");
// delete text file
$filename = "crc.txt";
unlink($filename);
// list files
list_dir($dir_handle,$path);
// tell them it's okay
echo 'Successfully dumped crcs. Safe to delete crc.php.';

function list_dir($dir_handle,$path)
{
    // open txt file for writing
    $myFile = "crc.txt";
    $fh = fopen($myFile, 'a') or die("Can't open file.");
    // print_r ($dir_handle);
    //running the while loop
    while (false !== ($file = readdir($dir_handle))) {
        $dir =$path.'/'.$file;
        if(is_dir($dir) && $file != '.' && $file !='..' )
        {
            $handle = @opendir($dir) or die("unable to open file $file");
            list_dir($handle, $dir);
        }elseif($file != '.' && $file !='..')
        {
               $stringData = "$dir,".getCRC($dir).",".filesize($dir)."|";
               fwrite($fh, $stringData);
        }
    }
    // Close the txt file
    fclose($fh);
    //closing the directory
    closedir($dir_handle);
   
}

function getCRC($dir)
{
    // Read the file into an array
    $data = file($dir);

    // Join the array into a string
    $data = implode('', $data);

    // Calculate the crc
    $crc = crc32($data);

    //convert from decimal to hexadecimal
    $crchex=DecHex($crc*1);

    //echo the result
    return "$crchex";
}
?>