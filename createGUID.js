/*
** Create a new GUID.
*/


function createGUID () {
	var scrtl = new ActiveXObject("Scriptlet.TypeLib");
	return scrtl.GUID;
}
