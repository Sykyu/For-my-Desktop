var ws = WScript.CreateObject("WScript.Shell");
ws.SendKeys( "% n" );  

// ��摜���쐬 �摜+���ԂŃt�@�C�����쐬
var jpegname = "�摜" + (new Date()).getTime() + ".jpeg";
var out = WScript.CreateObject("ADODB.Stream");
out.Type = 1;
out.Open();
out.SaveToFile( jpegname, 2 );
out.Close();

// MS�y�C���g���N���E�ő剻
var mspaint = ws.Run("mspaint.exe " + jpegname, 3);
WScript.sleep(250);

// WScript.Sleep( 1000 );
var ret = ws.AppActivate( mspaint );

// �摜��ۑ�
// �y�[�X�g
ws.SendKeys( "^v" ); 
// �ۑ�
ws.SendKeys( "^s" ); 
// �I��
WScript.sleep(250);
ws.SendKeys( "%{F4}" ); 