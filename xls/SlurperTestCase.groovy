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

class SlurperTestCase extends GroovyTestCase {
	def fileName = "temp-workbook.xls"
	def workbook
	def slurper
	protected void setUp(){
			workbook = new SimpleXlsBuilder().workbook{
			(1..2).each{
				sheet(name:"Sheet ${it}") {
					(1..3).each{row(0:"ZERO${it}",1:"ONE${it}",5:255615.453,7:false,10:new Date())}
				}
			}
		}
		workbook.saveToFile fileName
		slurper = new SimpleXlsSlurper(fileName)
	}
	protected void tearDown(){
		new File(fileName).delete()
	}
	void testTypesArePreserved() {
		assert slurper.sheets(0).rows(0).cells(0).'class' == String
		assert slurper.sheets(0).rows(0).cells(1).'class' == String
		def type = slurper.sheets(0).rows(0).cells(5).'class'
		assertTrue("Number type not preserved -> ${type}", Number.isAssignableFrom(type))
		assert slurper.sheets(0).rows(0).cells(7).'class' == Boolean
		assert slurper.sheets(0).rows(0).cells(10).'class' == Date
	}
	void testAddressesAreOk() {
		assertEquals(slurper.valueAt("Sheet 2!F1"),slurper.sheets("Sheet 2").rows(0).cells(5))
		assertEquals(slurper.sheets("Sheet 2").rows(0).cells(5), 255615.453)
	}
	void testIterationOk() {
		def sheetNum = 1
		slurper.each{sheet->
			assertEquals(sheet.name,"Sheet ${sheetNum++}")
			def rowNum = 0
			sheet.each{row->
				assertEquals(row.num,rowNum++)
				row.each{cellValue->
					assertNotNull(cellValue)
				}
			}
		}
		def rowNum = 0
		slurper.sheets("Sheet 2").each{row->
			assertEquals(row.num,rowNum++)
			row.each{cellValue->
				assertNotNull(cellValue)
			}
		}
		rowNum = 0
		slurper.sheets(1).each{row->
			assertEquals(row.num,rowNum++)
			row.each{cellValue->
				assertNotNull(cellValue)
			}
		}
		(0..1).each{i->
			(0..2).each{j->
				slurper.sheets(i).rows(j).each{cellValue->
					assertNotNull(cellValue)
				}
			}
		}
		slurper.sheets("Sheet 1").rows(1).each{cellValue->
			assertNotNull(cellValue)
		}
		shouldFail(IllegalArgumentException){slurper.sheets("404")}
		shouldFail(IllegalArgumentException){slurper.sheets(200)}
		shouldFail(IllegalArgumentException){slurper.sheets(0).rows(-1)}
		shouldFail(IllegalArgumentException){slurper.sheets(0).rows(0).cells(-1)}
		shouldFail(IllegalArgumentException){slurper.sheets(0).rows(0).cells(500)}
	}
}