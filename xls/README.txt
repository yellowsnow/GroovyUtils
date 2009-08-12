1- Introduction

SimpleXlsBuilder.groovy is a groovy builder for xls spreadsheets based on POI project (http://poi.apache.org)

2- Usage

2.1 System requirements

Groovy 1.6.4+ (and Java of course!)


2.2 Running the sample with Groovy / Grape environment


	grooy Demo.groovy


2.3 CLASSPATH for non Grape use

Please run the following command to get classpath dependencies : (check Groovy Grape usage for more options here : http://groovy.codehaus.org/Grape)

	grape resolve "org.apache.poi" "poi" "3.5-beta6" "org.apache.poi" "poi-ooxml" "3.5-beta6"

3- Licence

Copyright 2009 Yellow Snow 
Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with the License. You may obtain a copy of the License at 

	http://www.apache.org/licenses/LICENSE-2.0 

Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing permissions and limitations under the License. 