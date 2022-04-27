
*------------------------------------------------------------------------------*
* Table generation with putexcel
*------------------------------------------------------------------------------*

clear all

ssc install excelcol

webuse auto, clear

ta rep78
recode rep78 1/2=0 3=1 4/5=2, g(rep78_3)

* Categorical
tab1 foreign rep78_3
* Continuous
summarize price mpg headroom trunk weight length turn

* install excelcol - only need to install once, then comment
* ssc install excelcol

*------------------------------------------------------------------------------*
* Table 1. Descriptives - manual 
*------------------------------------------------------------------------------*

* Categorical variables 
tab1	foreign rep78_3

proportion foreign					// proportions command
eret list							// estimation results					
mat def prop = r(table)
mat list prop					// results

proportion foreign				// proportions command

di `e(N)'
di prop[1,1]
di prop[5,1]
di prop[6,1]

di prop[1,2]
di prop[5,2]
di prop[6,2]

* Setup excel file
putexcel set putexceltut_pres.xlsx, sheet(Table1_Manual) replace

proportion foreign				
mat def prop = r(table)
local n 	= `e(N)'			// store n
local p  	= prop[1,1]*100		// store prop
local ll 	= prop[5,1]*100		// store lower 95% ci
local ul 	= prop[6,1]*100		// store upper 95% ci
local ns  	= `n'*prop[1,1]		// store n for each level

putexcel B2 = "`ns'"			
putexcel C2 = "`p'%"		
putexcel D2 = "(`ll', `ul')"

*------------------------------------------------------------------------------*
* Basic loops 
*------------------------------------------------------------------------------*

* Categorical variables 
local cats	foreign rep78_3
			
foreach x of varlist `cats' {	
	
	prop `x'
	
}

* Counting			
forvalues i = 1/6 {	
	
	di 	`i' * 5
	
}

*------------------------------------------------------------------------------*
* Table 1. Descriptives - loops 
*------------------------------------------------------------------------------*

putexcel set putexceltut_pres.xlsx, sheet(Table1_Loops) modify

* Setup starting row/col
local row = 2 
local col = 1 
	
* Categorical variables 
local cats	foreign rep78_3
			
foreach x of varlist `cats' {		// loop over each variable in cats

	proportion `x' 					// proportions command
	mat def prop = r(table)			// store props as matrix
	local n = e(N)					// store n from prop command
	local ncols = colsof(prop)		// number of columns in prop matrix

	forvalues i = 1/`ncols' {		// loop columns 1 to ncols (number of columns)					
	
		local p  	= string(prop[1,`i']*100,"%9.0f")		// store prop
		local ll 	= string(prop[5,`i']*100,"%9.0f")		// store lower 95% ci
		local ul 	= string(prop[6,`i']*100,"%9.0f")		// store upper 95% ci
		local ns  	= string(`n'*prop[1,`i'],"%9.0f")		// store n for each level
			
		excelcol `col'
		putexcel `r(column)'`row' = "`x'"			// variable name
		local col = `col' + 1
		excelcol `col'
		putexcel `r(column)'`row' = "`ns'"			// n for each level
		local col = `col' + 1
		excelcol `col'
		putexcel `r(column)'`row' = "`p'%"			// prop
		local col = `col' + 1
		excelcol `col'
		putexcel `r(column)'`row' = "(`ll', `ul')"	// 95% CI
		local col = `col' + 1
	
		local row = `row' + 1	// go to next row
		local col = 1			// reset at column 1
	
	}
		
}

* Continuous variables 
local cons	price mpg headroom trunk weight length turn
	
local row = `row' + 1	// go to next row
foreach x of varlist `cons' {		// loop over variables in cons

		mean `x' 					// mean command
		local n = e(N)				// store n from mean command
		mat mean = r(table)			// store mean as matrix
		
		summ `x'									// command to get sd		
		local sd 	= string(r(sd),"%9.1f") 		// store sd
		
		local me 	= string(mean[1,1],"%9.1f")			// store mean
		local ll 	= string(mean[5,1],"%9.1f")			// store lower 95% ci
		local ul 	= string(mean[6,1],"%9.1f")			// store upper 95% ci
			
		excelcol `col'
		putexcel `r(column)'`row' = "`x'"			// variable name
		local col = `col' + 1
		excelcol `col'
		putexcel `r(column)'`row' = "`me'"			// mean
		local col = `col' + 1
		excelcol `col'
		putexcel `r(column)'`row' = "`sd'"			// sd
		local col = `col' + 1
		excelcol `col'
		putexcel `r(column)'`row' = "(`ll', `ul')"	// 95% ci
		local col = `col' + 1
		
		local row = `row' + 1	// go to next row
		local col = 1			// reset at column 1
	
}

*------------------------------------------------------------------------------*
* Table 2. Regressions - loops 
*------------------------------------------------------------------------------*

putexcel set putexceltut_pres.xlsx, sheet(Table2_Loops) modify

* Setup starting row/col
local row = 2 
local col = 1 

local model1 = " "
local model2 = "i.foreign i.rep78_3"

foreach mod in model1 model2 {
 
	foreach y of varlist price mpg {
		
		foreach x of varlist headroom trunk weight length turn {
	
			regress `y' c.`x' ``mod''
			mat def rtab = r(table)
				
			local i = 1	// row 1 for main first predictor
			
				local b  = string(rtab[1,`i'],"%9.2f")
				local ll = string(rtab[5,`i'],"%9.2f")
				local ul = string(rtab[6,`i'],"%9.2f")
				local p  = string(rtab[4,`i'],"%9.3f")
									
				excelcol `col'
				putexcel `r(column)'`row' = "`x'"
				local col2 = `col' + 1
				excelcol `col2'
				putexcel `r(column)'`row' = "`b'"
				local col2 = `col2' + 1			
				excelcol `col2'
				putexcel `r(column)'`row' = "(`ll', `ul')"
				local col2 = `col2' + 1			
				excelcol `col2'
				putexcel `r(column)'`row' = "`p'"		
				
				local row = `row' + 1
			
		}
			
	local row = `row' + 2
	
	}

	local col = `col' + 5
	local row = 2

}

exit
