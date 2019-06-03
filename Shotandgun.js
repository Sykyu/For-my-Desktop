var ws = WScript.CreateObject("WScript.Shell");
ws.SendKeys( "% n" );  

// 空画像を作成 画像+時間でファイル名作成
var jpegname = "画像" + (new Date()).getTime() + ".jpeg";
var out = WScript.CreateObject("ADODB.Stream");
out.Type = 1;
out.Open();
out.SaveToFile( jpegname, 2 );
out.Close();

// MSペイントを起動・最大化
var mspaint = ws.Run("mspaint.exe " + jpegname, 3);
WScript.sleep(250);

// WScript.Sleep( 1000 );
var ret = ws.AppActivate( mspaint );

// 画像を保存
// ペースト
ws.SendKeys( "^v" ); 
// 保存
ws.SendKeys( "^s" ); 
// 終了
WScript.sleep(250);
ws.SendKeys( "%{F4}" ); 