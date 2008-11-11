// the location of the ExcelDna.xll relative to the current directory
var relXll = "../../third-party/ExcelDna/Distribution/ExcelDna.xll";
// the location of the template for the .dna file relative to the current directory
// the template must have placeholders for $NAME and $SCRIPT
var relDnaTemplate = "makedna.dna";
var overwriteAllXlls = false;

var ForReading = 1;
var ForWriting = 2;
var File = WScript.CreateObject("Scripting.FileSystemObject");
// returns the current directory
function getCurrentDirectory() {
	return File.GetAbsolutePathName(".");
}
// returns the content of a file given by its path as a string
function read(fn) {
	var f = File.OpenTextFile(fn, ForReading);
	var text = f.ReadAll();
	f.close();
	return text;
}
// writes to a file given by its path the given text
function write(fn, text) {
	var f = File.OpenTextFile(fn, ForWriting, true);
	var text = f.Write(text);
	f.close();
}

var dir = getCurrentDirectory();
var xllFile = File.GetFile(File.buildpath(dir, relXll));
var dnaTemplate = read(File.buildpath(dir, relDnaTemplate));

function dnaFromTemplate(name, script) {
	return dnaTemplate.replace(/\$NAME/, name).replace(/\$SCRIPT/, script);
}
 
// returns an array of all file paths starting with "dna" in the given directory
function getDnaFiles(dir) {
	var res = new Array();
	for (var fc = new Enumerator(File.GetFolder(dir).files); !fc.atEnd(); fc.moveNext())
	{
		var f = fc.item();
		var fn = f.Name;
		if (fn.substr(0,3) == "dna") {
			res.push(fn);
		}
	}
	return res;
}
var files = getDnaFiles(dir);

// returns X for dnaX.fs
function getScriptName(fn) {
	return fn.substr(3,fn.length-6);
}
// generates the *.xll and *.dna files for the given file path
function generateDna(fn) {
	var name = getScriptName(fn);
	var fnXll = File.buildpath(dir, name + ".xll");
	var fnDna = File.buildpath(dir, name + ".dna");
	if (overwriteAllXlls || !File.FileExists(fnXll)) {
		xllFile.Copy(fnXll);
	}
	write(fnDna, dnaFromTemplate(name, read(fn)));
}

for (var i=0; i<files.length; i++) {
	var fn = files[i];
	WScript.Echo("Compiling " + fn);
	generateDna(fn);
}