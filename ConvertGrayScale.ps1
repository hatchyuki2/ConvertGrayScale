
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName presentationframework
Add-Type -AssemblyName System.Windows.Forms

# BitmapSourceに変換し，グレースケール化
function ConvertGrayScale($src_image) {

    $image_width = $src_image.Width
    $image_height = $src_image.Height

    $NewObject = New-Object System.Windows.Media.Imaging.FormatConvertedBitmap

    [System.Windows.Media.Imaging.BitmapSource]$bitmapSource

    # BitmapSourceに変換
    $bitmapSource = [System.Windows.Interop.Imaging]::CreateBitmapSourceFromHBitmap(
     $src_image.GetHbitmap(),
     [IntPtr]::Zero,
     [System.Windows.Int32Rect]::Empty,
     [System.Windows.Media.Imaging.BitmapSizeOptions]::FromEmptyOptions()
    )

    # グレースケール化
    $NewObject.BeginInit();
    $NewObject.Source = $bitmapSource
    $NewObject.DestinationFormat = [System.Windows.Media.PixelFormats]::Gray32Float
    $NewObject.EndInit()

    [System.Windows.Media.Imaging.BitmapSource]$grayScaleBitmap = $NewObject
    $outPutBitmap = New-Object System.Drawing.Bitmap(BitmapFromSource $grayScaleBitmap)

    return [System.Drawing.Bitmap]$outPutBitmap
}

# BitmapSourceをBitmapに変換
function BitmapFromSource([System.Windows.Media.Imaging.BitmapSource]$bitmapSource) {
    $ms = New-Object System.IO.MemoryStream
    $encoder = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
    
    $encoder.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create([System.Windows.Media.Imaging.BitmapSource]$bitmapSource))

    $encoder.Save($ms)

    $outBitmap = New-Object System.Drawing.Bitmap($ms)

    return $outBitmap
}


# 保存先のパス、保存先の親フォルダのパスを生成
$scriptPath = $MyInvocation.MyCommand.Path
$scriptFolder = Split-Path -Parent $scriptPath
$outFlolder = Join-Path $scriptFolder "output"

# 保存先のフォルダが存在しない場合にフォルダを自動生成
if (!(Test-Path $outFlolder))
{
     [Void][IO.Directory]::CreateDirectory($outFlolder)
}

$counter = 0

# 与えられたパス（c:\tmp\convert）から合致するファイルリストを再帰的に取得
Get-ChildItem |

# 取得したファイルを順番に処理
ForEach-Object {
    # 取得したオブジェクトがファイルの場合のみ処理（フォルダの場合はスキップ）
    if ($_.GetType().Name -eq "FileInfo" -And
        ($_.Extension -eq ".png" -or $_.Extension -eq ".jpg" ))
    {
        # 画像読み込み
        $_.Name
        $src_image = [System.Drawing.Image]::FromFile($_.FullName)
    
        # 画像処理の関数を実行
        $tmp_image = ConvertGrayScale $src_image
        
        $dst_image = New-Object System.Drawing.Bitmap($tmp_image[1])
        
        # 画像の保存
        $dst_image.Save($outFlolder + "\" + $_.BaseName + "_out.png", [System.Drawing.Imaging.ImageFormat]::Png)
    
        # オブジェクトを破棄
        $src_image.Dispose()
        $dst_image.Dispose()
    
        $counter++
    }
}

# メッセージボックスの表示
$Message = [String]$counter + "個のファイルを変換しました"
[System.Windows.Forms.MessageBox]::Show($Message, "変換結果") 
