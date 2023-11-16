/*
 (c) VNexsus 2023

 Licensed under the Apache License, Version 2.0 (the "License");
 you may not use this file except in compliance with the License.
 You may obtain a copy of the License at

     http://www.apache.org/licenses/LICENSE-2.0

 Unless required by applicable law or agreed to in writing, software
 distributed under the License is distributed on an "AS IS" BASIS,
 WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 See the License for the specific language governing permissions and
 limitations under the License.
 */
 
(function(window, undefined){
 
 
	window.Asc.plugin.init = function()  {
		$(document.body).addClass(window.Asc.plugin.getEditorTheme());
		var filepath = parent.AscDesktopEditor.LocalFileGetSourcePath();
		var editortype = window.Asc.plugin.info.editorType;
		var typepath;
		
		switch (window.Asc.plugin.info.editorType){
			case 'word': {
				typepath = 'word'
				break;
			}
			case 'cell': {
				typepath = 'xl'
				break;
			}
			case 'slide': {
				typepath = 'ppt'
				break;
			}
		}
		
		const req = new XMLHttpRequest();
		req.open("GET", filepath, true);
		req.responseType = "blob";

		req.onload = (event) => {
			const data = req.response;
			JSZip.loadAsync(data).then(function(zip) {
				var folder = zip.folder(typepath + '/embeddings');
				if(Object.values(folder.files).filter(function(file){return (file.name.indexOf(typepath + "/embeddings/") == 0 && file.dir == false)}).length > 0)
					folder.forEach(function(filename, file){
						file.async('uint8array').then(function(filedata){
							var e = new EMBFile(filename, filedata);
							$('#contents').append(e.generate());
						});
					});
				else
				    $('#contents').html('<div style="display: table;text-align: center;vertical-align: baseline;width: 100%;height: 100%;"><div style="display: table-cell;vertical-align: middle;">Документ не содержит вложенных файлов</div></div>');			
			});
		};
		req.send();
		
		var EMBFile = function(filename, filedata){
			var self = this;
			self.filename = filename;
			self.filedata = filedata;
			self.filetype = self.filename.substring(self.filename.lastIndexOf('.') + 1);
			
			self.generate = function(){
				var el = $('<div/>').addClass('item');
				var ico = $('<span/>').addClass('icon fiv-viv fiv-size-md fiv-icon-' + self.filetype);
				var fn = $('<span/>').addClass('name ellipsis').attr('data-tail','.' + self.filetype).text(self.filename.substring(0,self.filename.length - self.filetype.length - 1));
				el.append(ico);
				el.append(fn);
				el.on('click', function(){
					var blob=new Blob([self.filedata], {type: "application/pdf"});
					var link=document.createElement('a');
					link.href=window.URL.createObjectURL(blob);
					link.download=self.filename;
					link.click();
				});
				return el;
			}
			
			if(self.filetype == 'bin'){
				var ole = CFB.parse(filedata,{type: 'binary'});
				var container = Object.values(ole.FileIndex).find(function(val){ return val.name == '\1Ole10Native'});
				if(container){
					var pos = 0;
					var cbStream = reader.readUint(container.content, pos);
					pos+=4;
					var wDirectoryType = reader.readUshort(container.content, pos);
					pos+=2;
					var szFileNameWithExtension = reader.readString(container.content, pos);
					pos+=szFileNameWithExtension.length+1;
					var szFullyQualifiedFileName = reader.readString(container.content, pos);
					pos+=szFullyQualifiedFileName.length+1;
					var singleChar = container.content[pos];
					pos++;
					var singleChar2 = container.content[pos];
					pos++;
					var wDirectoryType2 = reader.readUshort(container.content, pos);
					pos+=2;
					var lengthOfString = reader.readUint(container.content, pos);
					pos+=4;
					var szFullyQualifiedFileName2 = reader.readString(container.content, pos);
					pos+=szFullyQualifiedFileName2.length+1;
					var attachmentSize = reader.readUint(container.content, pos);
					pos+=4;
					self.filename = szFileNameWithExtension;
					self.filetype = self.filename.substring(self.filename.lastIndexOf('.') + 1);
					self.filedata = new Uint8Array(container.content.splice(pos));
				}
				else if(Object.values(ole.FileIndex).find(function(val){ return val.name == 'CONTENTS'})){
					container = Object.values(ole.FileIndex).find(function(val){ return val.name == 'CONTENTS'});
					var start, end;
					for(var pos = 0; pos < container.size-4; pos++){
						if(String.fromCharCode(container.content[pos], container.content[pos+1], container.content[pos+2], container.content[pos+3]) == "%PDF"){
							start = pos;
							break;
						}
					}
					if(typeof start !== 'undefined'){
						for(var pos = container.size-5; pos > start ; pos--){
							if(String.fromCharCode(container.content[pos], container.content[pos+1], container.content[pos+2], container.content[pos+3], container.content[pos+4]) == "%%EOF"){
								end = pos+5;
								break;
							}
						}
					}
					if(typeof start !== 'undefined' && typeof end !== 'undefined'){
						self.filename = "Документ Adobe Acrobat.pdf";
						self.filetype = self.filename.substring(self.filename.lastIndexOf('.') + 1);
						self.filedata = new Uint8Array(container.content.splice(start, end));
					}
				}
			}
			
		}

		var reader = {
			uint8 : new Uint8Array(4),
			readShort	: function(buff,p)  {  var u8=reader.uint8;  u8[0]=buff[p];  u8[1]=buff[p+1];  return reader.int16 [0];  },
			readUshort	: function(buff,p)  {  var u8=reader.uint8;  u8[0]=buff[p];  u8[1]=buff[p+1];  return reader.uint16[0];  },
			readInt		: function(buff,p)  {  var u8=reader.uint8;  u8[0]=buff[p];  u8[1]=buff[p+1];  u8[2]=buff[p+2];  u8[3]=buff[p+3];  return reader.int32 [0];  },
			readUint	: function(buff,p)  {  var u8=reader.uint8;  u8[0]=buff[p];  u8[1]=buff[p+1];  u8[2]=buff[p+2];  u8[3]=buff[p+3];  return reader.uint32[0];  },
			readFloat	: function(buff,p)  {  var u8=reader.uint8;  u8[0]=buff[p];  u8[1]=buff[p+1];  u8[2]=buff[p+2];  u8[3]=buff[p+3];  return reader.float32[0];  },
			readASCII	: function(buff,p,l){  var s = "";  for(var i=0; i<l; i++) s += String.fromCharCode(buff[p+i]);  return s;    },
			readString	: function(buff,p)	{  var e = buff.indexOf(0,p); return new TextDecoder("windows-1251").decode(new Uint8Array(buff.slice(p,e)));}
		}

		reader.int16  = new Int16Array (reader.uint8.buffer);
		reader.uint16 = new Uint16Array(reader.uint8.buffer);
		reader.int32  = new Int32Array (reader.uint8.buffer);
		reader.uint32 = new Uint32Array(reader.uint8.buffer);
		reader.float32 = new Float32Array(reader.uint8.buffer);		
		
	};
	
	window.Asc.plugin.getEditorTheme = function(){
		if(window.localStorage.getItem("ui-theme-id")){
			var match = window.localStorage.getItem("ui-theme-id").match(/\S+\-(\S+)/);
			if(match.length==2)
				return "theme-" + match[1];
		}
		return "theme-light";
	}
	
	window.Asc.plugin.onThemeChanged = function(theme){
		window.Asc.plugin.onThemeChangedBase(theme);
		$(document.body).removeClass("theme-dark theme-light").addClass(window.Asc.plugin.getEditorTheme());
	}

    window.Asc.plugin.button = function(id)  {
		this.executeCommand("close", "");
    };


})(window, undefined);
