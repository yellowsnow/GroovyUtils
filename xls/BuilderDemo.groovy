/*
Copyright 2009 Yellow Snow 

Licensed under the Apache License, Version 2.0 (the "License"); you may not 
use this file except in compliance with the License. You may obtain a copy of 
the License at 

	http://www.apache.org/licenses/LICENSE-2.0 

Unless required by applicable law or agreed to in writing, software 
distributed under the License is distributed on an "AS IS" BASIS, WITHOUT 
WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the 
License for the specific language governing permissions and limitations under 
the License. 
*/

@Grapes([
	@Grab(group='org.apache.poi', module='poi', version='3.5-beta6'),
	@Grab(group='org.apache.poi', module='poi-ooxml', version='3.5-beta6')
])

class BuilderDemo {
	static main(args) {
		def workbook = new SimpleXlsBuilder().workbook(templateFileName:"template.xls"/*OPTIONAL, can also use InputStream via the 'templateInputStream' argument*/){
			sheet(name:"Feuil1") {
				(1..3).each{row(0:"ZERO${it}",1:"ONE${it}",5:2556,6:-25888,7:898956,10:new Date())}
			}
			sheet(name:"Feuil2") {
				(1..3).each{row(0:"ZERO${it}",1:"ONE${it}",5:6,6:-25888,7:898956,10:new Date())}
			}
		}
		def fileName = "workbook.xls"
		workbook.saveToFile fileName
		//Launches the file in the spreadsheet
		def osname = System.getProperty("os.name")
		if (osname.toLowerCase().contains("win")) {
			"rundll32 url.dll,FileProtocolHandler $fileName".execute()
		} else if (osname.toLowerCase().contains("mac")) {
			"open $fileName".execute()
		} else {
			// sorry, i suppose it's gnome on X :-)
			"open $fileName".execute()
		}
	}
}