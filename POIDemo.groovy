def workbook = new SimpleXlsBuilder().workbook(templateFileName:"template.xls"){
	sheet(name:"Feuil1") {
		cell(ref:"C10",value:"test 1 :")
		cell(ref:"C11",value:"test 2 :")
		cell(ref:"C12",value:"test 2 :")
	}
	(3..6).each{cell(ref:"My New Sheet!C${it}",value:"NONE")}
}
def fileName = "workbook.xls"
workbook.saveToFile fileName
"rundll32 url.dll,FileProtocolHandler $fileName".execute()
