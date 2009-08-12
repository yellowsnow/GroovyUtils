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

class SlurperDemo {
	static main(args) {
		def slurper = new SimpleXlsSlurper("workbook.xls")
		println "C10: ${slurper.valueAt("C10")}"
		println "C11: ${slurper.valueAt("C11")}"
		println "My New Sheet!C3: ${slurper.sheet("My New Sheet").valueAt("C3")}"
		println "My New Sheet!C4: ${slurper.valueAt("My New Sheet!C4")}"
	}
}