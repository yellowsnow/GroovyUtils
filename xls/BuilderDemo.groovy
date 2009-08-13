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
		/*templateFileName is OPTIONAL, one can also use a template from an InputStream via the 'templateInputStream' argument*/
		def workbook = new SimpleXlsBuilder().workbook(templateFileName:"template.xls"){
			sheet(name:"Feuil1") {
				(1..3).each{row(0:"ZERO${it}",1:"ONE${it}",5:2556,6:-25888,7:898956,10:new Date())}
			}
			sheet(name:"Feuil2") {
				(1..3).each{row(0:"ZERO${it}",1:"ONE${it}",5:6,6:-25888,7:898956,10:new Date())}
			}
			sheet(name:"Feuil3") {
				row("A25":"TEST FREE 1","B25":new Date(),"C25":new Date())
				row("A26":"TEST FREE 2","B26":new Date(),"C26":new Date())
				row {
					cell(value:"1")
					cell(value:"2")
					cell(value:"3")
					cell(value:"4")
				}
				cell(ref:"C10",value:"test 1 :")
				cell(ref:"C11",value:"test 2 :")
				cell(ref:"C12",value:"test 2 :")
				row {
					cell(value:"test 4")
					cell(x:10, value:"test 5")
				}
				row(y:20) {
					cell(value:"test 6")
					cell(x:10, value:"test 7")
				}
			}
			(3..6).each{cell(ref:"My New Sheet!C${it}",value:"NONE")}
			sheet(name:"Last") {
				row {
					cell(value:"1")
					cell(value:"2")
					cell(value:"3")
					cell(value:"4")
				}
				row("A25":"LAST TEST FREE 1","B25":new Date(),"C25":new Date())
				row("A26":"LAST TEST FREE 2","B26":new Date(),"C26":new Date())
				row {
					cell(value:10)
					cell(value:20)
					cell(value:30)
					cell(value:40)
				}
				row(0:"ZERO",1:"ONE",5:"FIVE",10:"TEN")
				row(0:"2ZERO",1:"2ONE",5:"2FIVE",10:"2TEN")
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