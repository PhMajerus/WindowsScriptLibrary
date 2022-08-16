/*
** Generates a new GUID.
** 
** Dependencies:
** - Windows Script Component Runtime
**   (part of the base Windows Script platform)
*/


function createGuid () {
	var scrtl = new ActiveXObject("Scriptlet.TypeLib");
	return scrtl.GUID.substring(0,38);
}
