/*************************************************
****** Consolidate and check updates of tracker***
****** Date: January 2020            *************
****** mbidinlib@poverty-action.org       ********/

* This do file consolidates 5 trackers from 5 mobilizers working on the fied
* It also creates an excel output file which contains an enumerator dashboard



* Set paths
loc dropbox  "dropox" // dropdox path
loc fname   "Tracker_"  // name of excel tracker file
loc outfolder "folder"		// output folder

cap mkdir "`outfolder'/`c(current_date)'"
cap mkdir "`outfolder'/overall"
loc dfolder  "`outfolder'/`c(current_date)'"



* import individual trackers
forval j = 1/20 {

	import excel using "`dropbox'/`fname'`j'/`fname'`j'.xlsx", firstrow clear
	drop A
	drop if _n <5 | (missing(q01) & missing(q05))
	
	save "`dfolder'/`fname'`j'" , replace   // Save a invividual copy for reference
}


noi di "Hey I changed the do file like this"


// consolidate trackers
use "`dfolder'/`fname'1", clear

forval j = 2/5 {
	append using "`dfolder'/`fname'`j'"	
	save "`outfolder'/overall/Overall_Tracker_`c(current_date)'", replace
}


* Export duplicates as a table in excel
duplicates t q05, gen(dt)
bys q05 dt: gen cnt =_n
keep if cnt ==1 & dt>1

cd "`outfolder'/overall"
local fname1 "Tracker_Analisys.xlsx"
putexcel set Tracker_Analisys.xlsx, modify
loc fcol = 3
loc lcol = 8

loc ln = _N
forval j = 1/`ln' {
	loc v = `j' + 2
	putexcel C`v' = q05[`j'] //School_ID
	putexcel D`v' = q04[`j'] // Name of School
	putexcel E`v' = q06[`j'] // Locality
	putexcel F`v' = q07[`j'] //District
	putexcel G`v' = q08[`j'] // Type
	putexcel H`v' = dt[`j']-1 // Number of Duplicates
}
putexcel C2 = "School_ID"
putexcel D2 = "Name of School"
putexcel E2 = "Locality"
putexcel F2 = "District"
putexcel G2 = "Type"
putexcel H2 = "Number extra copies/Duplicates"
putexcel E1:G1 = "Duplicates", merge bold

* call mata to format the output
mata : format_dups()


// Export and create an enumerator dashboard

use "`outfolder'/overall/Overall_Tracker_`c(current_date)'", clear
bys q01: gen num2 = _n
by  q01: egen max2 = max(num2)
keep if num2  == 1
loc fcol = 11
loc lcol = 13
loc ln = _N


forval j = 1/`ln' {
	loc v = `j' + 2
	
	putexcel K`v' = q01[`j'] // Enum ID
	putexcel L`v' = q02[`j'] // Enum Name
	putexcel M`v' = max2[`j'] // Number of Surveys
}
putexcel K2 = "Enum ID"
putexcel L2 = "Enum Name"
putexcel M2 = "Number of Surveys"
putexcel K1:M1 = "Enumerator Dashboard", merge bold

* call mata to format enumerator dashboard
mata : format_dups()



// Formatig the outputs in mata

mata : 
mata clear
void format_dups(){

	 filename = st_local("fname1")
	 number = st_local("ln")
	 sheet = "Sheet1"
	 fcol = st_local("fcol")
	 lcol = st_local("lcol")
	 
	 class xl scalar b
	 
	 dp = strtoreal(number)
	 fc = strtoreal(fcol)
	 lc = strtoreal(lcol)
	 bd = dp + 2
	
	b.load_book(filename)
	b.set_sheet(sheet)
	
	b.set_sheet_gridlines(sheet,"off")
	b.set_border((2, bd), (fc, lc), "thin")
	b.set_top_border((2, 2),  (fc, lc), "thick")
	b.set_bottom_border((2, 2), (fc, lc), "thick")
	b.set_left_border((2, bd), fc, "thick")
	b.set_right_border((2, bd), lc, "thick")
	b.set_bottom_border(bd, (fc, lc), "thick")

	b.close_book()
	
	b.set_column_width(3, lc - 1 , 12)
	b.set_column_width(lc, lc, 12)
	
}
end







