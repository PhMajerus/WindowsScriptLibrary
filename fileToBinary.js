/*
** Create a Binary buffer (byte array / VT_ARRAY|VT_UI1) from a file path,
** using only components included with Windows.
** 
** Dependencies:
** - Microsoft ActiveX Data Objects ("ADODB.Stream")
*/


function fileToBinary (filepath) {
	var stream = new ActiveXObject("ADODB.Stream");
	stream.type = 1 /* adTypeBinary */;
	stream.open();
	stream.loadFromFile(filepath);
	return stream.read();
}
