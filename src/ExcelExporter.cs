using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Excel;

public class ExcelExporter{
	
	private class Field{
		public int index;
		public bool isMap;
		public bool isList;
		public string name;
		public string desc;
		public string type;
		public string mapKeyType;
		public string mapValueType;
		public string listValueType;
	}
	
	private class Table{
		public string name;
		public bool isConfig;
		public List<List<string>> lines;
		public List<Field> fields;
		public List<Field> keys;
	}
	
	private const int START=2;
	private const int NAME_ROW=0;
	private const int DESC_ROW=1;
	private const int CTRL_ROW=2;
	private const int TYPE_ROW=3;
	private const int DATA_ROW=4;
	
	public static void Export(string inputFolder,string outputFolder){
		Dictionary<string,string> files=new Dictionary<string,string>();
		CollectExcel(files,Path.Combine(inputFolder,"client"));
		CollectExcel(files,Path.Combine(inputFolder,"global"));
		foreach(var pair in files){
			var name=pair.Key;
			var file=pair.Value;
			if(File.Exists(Path.Combine(Path.GetDirectoryName(file),"~$"+name+".xlsx"))){
				var tempFile=Path.GetTempFileName();
				File.Copy(file,tempFile,true);
				ExportExcel(name,tempFile,outputFolder);
				File.Delete(tempFile);
			}
			else{
				ExportExcel(name,file,outputFolder);
			}
		}
	}
	
	private static void CollectExcel(Dictionary<string,string> files,string folder){
		foreach(var file in Directory.GetFiles(folder,"*.xlsx")){
			var fileName=Path.GetFileNameWithoutExtension(file);
			if(fileName.IndexOf('~')<0){
				files.Add(fileName,file);
			}
		}
	}
	
	private static void ExportExcel(string name,string inputFile,string outputFolder){
		var lines=new List<List<string>>();
		bool isConfig=name.Contains("Config");
		using(var stream=File.Open(inputFile,FileMode.Open,FileAccess.ReadWrite)){
			using(var excelReader=ExcelReaderFactory.CreateOpenXmlReader(stream)){
				var result=excelReader.AsDataSet();
				var sheet=result.Tables[0];
				var rows=sheet.Rows;
				var rowCount=rows.Count;
				var columnCount=sheet.Columns.Count;
				var dataStartIndex=START+DATA_ROW;
				if(isConfig){
					for(int lineIndex=START;lineIndex<columnCount&&lineIndex<=dataStartIndex;lineIndex++){
						var line=new List<string>();
						for(int fieldIndex=START;fieldIndex<rowCount;fieldIndex++){
							line.Add(rows[fieldIndex][lineIndex].ToString());
						}
						lines.Add(line);
					}
				}
				else{
					for(int lineIndex=START;lineIndex<rowCount;lineIndex++){
						var line=new List<string>();
						var sheetLine=rows[lineIndex];
						var isEmptyLine=lineIndex>dataStartIndex;
						for(int fieldIndex=START;fieldIndex<columnCount;fieldIndex++){
							var text=sheetLine[fieldIndex].ToString();
							line.Add(text);
							if(isEmptyLine){
								if(!string.IsNullOrWhiteSpace(text)){
									isEmptyLine=false;
								}
							}
						}
						if(!isEmptyLine){
							lines.Add(line);
						}
					}
				}
			}
		}
		if(lines.Count>=DATA_ROW){
			var table=new Table();
			table.name=name;
			table.lines=lines;
			table.isConfig=isConfig;
			ParseTable(table);
			if(CheckTable(table)){
				ExportTable(table,Path.Combine(outputFolder,name+".cs"));
			}
		}
	}
	
	private static HashSet<string> keywards=new HashSet<string>(){
		"abstract","as","base","bool","break","byte","case","catch","char","checked",
		"class","const","continue","decimal","default","delegate","do","double","else","enum",
		"event","explicit","extern","false","finally","static","float","for","foreach","goto",
		"if","implicit","in","int","interface","internal","is","lock","long","namespace",
		"new","null","object","operator","out","override","params","private","protected","public",
		"readonly","ref","return","sbyte","sealed","short","sizeof","stackalloc","static","string",
		"struct","switch","this","throw","true","try","typeof","uint","ulong","unchecked",
		"unsafe","ushort","using","virtual","void","volatile","while","all","i",
	};
	
	private static Dictionary<string,string> types=new Dictionary<string,string>(){
		["bool"]="bool",["boolean"]="bool",
		["byte"]="byte",["uint8"]="byte",
		["sbyte"]="sbyte",["int8"]="sbyte",
		["short"]="short",["int16"]="short",
		["ushort"]="ushort",["uint16"]="ushort",
		["int"]="int",["rune"]="int",["int32"]="int",
		["uint"]="uint",["uint32"]="uint",
		["long"]="long",["int64"]="long",
		["ulong"]="ulong",["uint64"]="ulong",
		["float"]="float",["float32"]="float",
		["double"]="double",["float64"]="double",
		["string"]="string",["lang"]="string",
	};
	
	private static string ConvertType(string type){
		if(types.TryGetValue(type.Trim().ToLower(),out var value)){
			return value;
		}
		return null;
	}
	
	private static void ParseTable(Table table){
		table.fields=new List<Field>();
		table.keys=new List<Field>();
		var nameRow=table.lines[NAME_ROW];
		var descRow=table.lines[DESC_ROW];
		var ctrlRow=table.lines[CTRL_ROW];
		var typeRow=table.lines[TYPE_ROW];
		var fieldCount=typeRow.Count;
		for(var c=0;c<fieldCount;c++){
			var name=nameRow[c].Trim();
			var desc=descRow[c].Trim();
			var ctrl=ctrlRow[c].Trim();
			var type=typeRow[c].Trim();
			if(!string.IsNullOrEmpty(name)&&!string.IsNullOrEmpty(type)&&!ctrl.StartsWith("S")){
				var field=new Field();
				field.index=c;
				if(ctrl.StartsWith("K")){
					table.keys.Add(field);
				}
				field.name=name;
				field.desc=desc;
				if(type.StartsWith("[]")){
					field.isList=true;
					field.listValueType=ConvertType(type.Substring(2));
					field.type=field.listValueType+"[]";
					if(field.listValueType!=null){
						table.fields.Add(field);
					}
				}
				else if(type.StartsWith("map[")){
					field.isMap=true;
					var keyIndex=type.LastIndexOf(']')+1;
					if(keyIndex>5){
						field.mapKeyType=ConvertType(type.Substring(4,keyIndex-5));
						field.mapValueType=ConvertType(type.Substring(keyIndex,type.Length-keyIndex));
						field.type="Dictionary<"+field.mapKeyType+","+field.mapValueType+">";
						if(field.mapKeyType!=null&&field.mapValueType!=null){
							table.fields.Add(field);
						}
					}
				}
				else{
					field.type=ConvertType(type);
					if(field.type!=null){
						table.fields.Add(field);
					}
				}
			}
		}
		if(!table.isConfig&&table.keys.Count==0&&table.fields.Count>0){
			table.keys.Add(table.fields[0]);
		}
	}
	
	private static bool CheckTable(Table table){
		if(table.fields.Count<0){
			return false;
		}
		foreach(var field in table.fields){
			if(keywards.Contains(field.name)){
				return false;
			}
		}
		return true;
	}
	
	private const string EOL="\r\n";
	private static void ExportTable(Table table,string file){
		var builder=new StringBuilder();
		builder.Append($"using System.Collections.Generic;").Append(EOL);
		builder.Append($"namespace Excel{{").Append(EOL);
		builder.Append($"	public class {table.name}{{").Append(EOL);
		if(table.isConfig){
			AppendConfig(builder,table);
		}
		else{
			AppendTable(builder,table);
		}
		builder.Append($"	}}").Append(EOL);
		builder.Append($"}}");
		File.WriteAllText(file,builder.ToString());
		builder.Clear();
	}
	
	private static void AppendConfig(StringBuilder builder,Table table){
		var rowCount=table.lines.Count-DATA_ROW;
		foreach(var field in table.fields){
			builder.Append($"		").Append(EOL);
			builder.Append($"		/** {field.desc} */").Append(EOL);
			builder.Append($"		public readonly static {field.type} {field.name}=");
			AppendField(builder,field,table.lines[DATA_ROW][field.index]);
			builder.Append($";").Append(EOL);
		}
	}
	
	private static void AppendTable(StringBuilder builder,Table table){
		var rowCount=table.lines.Count-DATA_ROW;
		var keyCount=table.keys.Count;
		var keyNames=new string[keyCount+1];
		var keyTypes=new string[keyCount+1];
		var className=table.name;
		keyNames[keyCount]=className;
		keyTypes[keyCount]=className;
		for(var i=0;i<keyCount;i++){
			var key=table.keys[i];
			keyNames[i]=key.name;
			keyTypes[i]=key.type;
		}
		builder.Append($"		").Append(EOL);
		builder.Append($"		private static {GetDictType(keyTypes,0)} {GetDictName(keyNames,0)};").Append(EOL);
		builder.Append($"		private static void Add(");
		for(var i=0;i<keyCount;i++){
			builder.Append($@"{keyTypes[i]} {keyNames[i]},");
		}
		builder.Append($@"int i){{").Append(EOL);
		for(var i=1;i<keyCount;i++){
			builder.Append($"			if(!{GetDictName(keyNames,i-1)}.TryGetValue({keyNames[i-1]},out var {GetDictName(keyNames,i)})){{").Append(EOL);
			builder.Append($"				{GetDictName(keyNames,i)}=new {GetDictType(keyTypes,i)}();").Append(EOL);
			builder.Append($"				{GetDictName(keyNames,i-1)}.Add({keyNames[i-1]},{GetDictName(keyNames,i)});").Append(EOL);
			builder.Append($"			}}").Append(EOL);
		}
		builder.Append($"			{GetDictName(keyNames,keyCount-1)}[{keyNames[keyCount-1]}]=new {className}(i);").Append(EOL);
		builder.Append($@"		}}").Append(EOL);
		
		builder.Append($"		public static {className} Get(");
		for(var i=0;i<keyCount;i++){
			if(i>0){
				builder.Append(',');
			}
			builder.Append($"{keyTypes[i]} {keyNames[i]}");
		}
		builder.Append($"){{").Append(EOL);
		builder.Append($"			if({GetDictName(keyNames,0)}==null){{").Append(EOL);
		builder.Append($"				{GetDictName(keyNames,0)}=new {GetDictType(keyTypes,0)}();").Append(EOL);
		builder.Append($"				for(int i=0;i<{rowCount};i++){{").Append(EOL);
		builder.Append($"					Add(");
		for(var i=0;i<keyCount;i++){
			builder.Append($"i_{keyNames[i]}[i],");
		}
		builder.Append($"i);").Append(EOL);
		builder.Append($"				}}").Append(EOL);
		builder.Append($"			}}").Append(EOL);
		for(var i=1;i<keyCount;i++){
			builder.Append($"			if(!{GetDictName(keyNames,i-1)}.TryGetValue({keyNames[i-1]},out var {GetDictName(keyNames,i)})){{").Append(EOL);
			builder.Append($"				return null;").Append(EOL);
			builder.Append($"			}}").Append(EOL);
		}
		builder.Append($"			{GetDictName(keyNames,keyCount-1)}.TryGetValue({keyNames[keyCount-1]},out var {keyNames[keyCount]});").Append(EOL);
		builder.Append($"			return {keyNames[keyCount]};").Append(EOL);
		builder.Append($"		}}").Append(EOL);
		builder.Append($"		").Append(EOL);
		builder.Append($"		private static {className}[] all;").Append(EOL);
		builder.Append($"		public static {className}[] GetAll(){{").Append(EOL);
		builder.Append($"			if(all==null){{").Append(EOL);
		builder.Append($"				all=new {className}[{rowCount}];").Append(EOL);
		builder.Append($"				for(int i=0;i<{rowCount};i++){{").Append(EOL);
		builder.Append($"					all[i]=Get(");
		for(var i=0;i<keyCount;i++){
			if(i>0){
				builder.Append(',');
			}
			builder.Append($"i_{table.keys[i].name}[i]");
		}
		builder.Append($");").Append(EOL);
		builder.Append($"				}}").Append(EOL);
		builder.Append($"			}}").Append(EOL);
		builder.Append($"			return all;").Append(EOL);
		builder.Append($"		}}").Append(EOL);
		builder.Append($"		").Append(EOL);
		builder.Append($"		private int i;").Append(EOL);
		builder.Append($"		private {className}(int i){{").Append(EOL);
		builder.Append($"			this.i=i;").Append(EOL);
		builder.Append($"		}}").Append(EOL);
		foreach(var field in table.fields){
			var fieldType=field.type;
			var fieldName=field.name;
			builder.Append($"		").Append(EOL);
			builder.Append($"		private static {fieldType}[] i_{fieldName};").Append(EOL);
			builder.Append($"		/** {field.desc} */").Append(EOL);
			builder.Append($"		public {fieldType} {fieldName}{{").Append(EOL);
			builder.Append($"			get{{").Append(EOL);
			builder.Append($"				if(i_{fieldName}==null){{").Append(EOL);
			builder.Append($"					i_{fieldName}=new {fieldType}[]{{");
			for(var i=0;i<rowCount;i++){
				if(i>0){
					builder.Append(',');
				}
				var fieldData=table.lines[DATA_ROW+i][field.index];
				AppendField(builder,field,fieldData);
			}
			builder.Append($"}};").Append(EOL);
			builder.Append($"				}}").Append(EOL);
			builder.Append($"				return i_{fieldName}[i];").Append(EOL);
			builder.Append($"			}}").Append(EOL);
			builder.Append($"		}}").Append(EOL);
		}
	}

	private static void AppendField(StringBuilder builder,Field field,string data){
		if(field.isList||field.isMap){
			builder.Append($"new {field.type}{{");
			if(!string.IsNullOrEmpty(data)){
				var items=data.Split(',');
				for(var i=0;i<items.Length;i++){
					if(i>0){
						builder.Append(',');
					}
					var item=items[i];
					if(field.isList){
						AppendData(builder,field.listValueType,item);
					}
					else{
						var index=item.IndexOf(':');
						var key=item.Substring(0,index<0?item.Length:index);
						var value=index<0?"":item.Substring(index+1,item.Length-index-1);
						builder.Append('[');
						AppendData(builder,field.mapKeyType,key);
						builder.Append(']');
						builder.Append('=');
						AppendData(builder,field.mapValueType,value);
					}
				}
			}
			builder.Append('}');
		}
		else{
			AppendData(builder,field.type,data);
		}
	}
	
	private static void AppendData(StringBuilder builder,string type,string data){
		if(string.IsNullOrEmpty(data)){
			if(type=="string"){
				builder.Append("\"\"");
			}
			else if(type=="bool"){
				builder.Append("false");
			}
			else{
				builder.Append("0");
			}
		}
		else{
			if(type=="int"){
				int.TryParse(data,out var value);
				builder.Append(value.ToString());
			}
			else if(type=="string"){
				builder.Append('"');
				foreach(var c in data){
					if(c=='"'){
						builder.Append("\\\"");
					}
					else if(c=='\t'){
						builder.Append("\\t");
					}
					else if(c=='\r'){
						builder.Append("\\r");
					}
					else if(c=='\n'){
						builder.Append("\\n");
					}
					else if(c=='\\'){
						builder.Append("\\\\");
					}
					else if(c<32){
						builder.Append("\\u");
						builder.Append(((int)c).ToString("X4"));
					}
					else{
						builder.Append(c);
					}
				}
				builder.Append('"');
			}
			else if(type=="float"){
				float.TryParse(data,out var value);
				builder.Append(value.ToString());
				builder.Append('f');
			}
			else if(type=="bool"){
				if(data=="1"){
					builder.Append("true");
				}
				if(data=="0"){
					builder.Append("false");
				}
				else{
					bool.TryParse(data,out var value);
					builder.Append(value?"true":"false");
				}
			}
			else if(type=="byte"){
				byte.TryParse(data,out var value);
				builder.Append(value.ToString());
			}
			else if(type=="sbyte"){
				sbyte.TryParse(data,out var value);
				builder.Append(value.ToString());
			}
			else if(type=="short"){
				short.TryParse(data,out var value);
				builder.Append(value.ToString());
			}
			else if(type=="ushort"){
				ushort.TryParse(data,out var value);
				builder.Append(value.ToString());
			}
			else if(type=="uint"){
				uint.TryParse(data,out var value);
				builder.Append(value.ToString());
			}
			else if(type=="long"){
				long.TryParse(data,out var value);
				builder.Append(value.ToString());
			}
			else if(type=="ulong"){
				ulong.TryParse(data,out var value);
				builder.Append(value.ToString());
			}
			else if(type=="double"){
				double.TryParse(data,out var value);
				builder.Append(value.ToString());
			}
		}
	}
	
	private static string GetDictName(string[] names,int index){
		return string.Join("_",names,index,names.Length-index);
	}
	
	private static string GetDictType(string[] types,int index){
		var builder=new StringBuilder();
		AppendDictType(builder,types,index);
		return builder.ToString();
	}
	
	private static void AppendDictType(StringBuilder builder,string[] types,int index){
		if(index<types.Length-1){
			builder.Append("Dictionary<");
			builder.Append(types[index]);
			builder.Append(',');
			AppendDictType(builder,types,index+1);
			builder.Append('>');
		}
		else{
			builder.Append(types[index]);
		}
	}
}