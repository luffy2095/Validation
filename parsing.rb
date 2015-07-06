#!/usr/bin/ruby
require 'spreadsheet'

class Automation

############masterKey arrays#############

@@master_name=Array.new
@@master_item=Array.new
@@master_godown=Array.new

############Test arrays#############

@@test_name=Array.new
@@test_item=Array.new
@@test_godown=Array.new

#############Returned_values Arrays###############

@@error_array_index=0 ####its basically no of error
@@error_array=Array.new  ####error row number
@@error_array_type=Array.new    ####error type

############Numbers############################

@@n_of_checks=0
@@n_of_master=0


#####################################################

def parsing_master(f)
	no_of_master_parsed=0
	book1 = Spreadsheet.open(f,'rb')
	sheet = book1.worksheet 0 ##parse each worksheet
	
	sheet.each 1 do |row|
		if row[0]==nil 
		break
		end
		@@master_name[no_of_master_parsed]=row[0]
		@@master_item[no_of_master_parsed]=row[1]
		@@master_godown[no_of_master_parsed]=row[2]
		no_of_master_parsed+=1
	end		###end of row parsing
	##end   ###end of sheet parsing
	
	@@n_of_master=no_of_master_parsed
	#puts no_of_master_parsed
		
end		#####end of parsing_master
  ##############end of class

#########################################################

def parsing_checks(f)
	no_of_test_parsed=0
	book2 = Spreadsheet.open(f,'rb')
	sheet = book2.worksheet 0 ##parse each worksheet
	
	sheet.each 1 do |row|
	
		if row[4]==nil 
		break
		end
		@@test_name[no_of_test_parsed]=row[4]
		@@test_item[no_of_test_parsed]=row[5]
		@@test_godown[no_of_test_parsed]=row[20]
		no_of_test_parsed+=1
	end		###end of row parsing
	##end   ###end of sheet parsing
	@@n_of_checks=no_of_test_parsed
	
	#@@test_name.collect!{|x| x.rstrip}
	#@@test_item.collect!{|x| x.rstrip}
	#@@test_godown.collect!{|x| x.rstrip}

	#puts no_of_test_parsed
		##puts no_of_test_parsed
		
end		#####end of parsing_checks
############################################################

def error_hilighting

	no_of_errors=(@@error_array_index)
	
	#book3=Spreadsheet.open(f,'rb') ###opening for editing
	#sheet = book3.worksheet 0
	
	
	book4=Spreadsheet::Workbook.new   ##creating new sheet
	sheet_new=book4.create_worksheet
	
	sheet_new.column(0).width=40
	sheet_new.column(1).width=50
	
	
	format = Spreadsheet::Format.new    :color => :blue,
						:weight => :bold,
						:size => 20,
						:align => :center
						
	format1 = Spreadsheet::Format.new   :align => :center
	i=0
	if no_of_errors==0
		sheet_new.row(0).default_format = format ##changing format
		sheet_new.column(0).default_format=format1
		sheet_new.column(1).default_format=format1		
		#row1=sheet.row(row_no)
		title=sheet_new.row(0)
		title[0]='Row Numbers'
		title[1]='Error Type'
		error_type='There are no errors in the UDI file.All entries are valid'
		sheet_new[i+2,1]=error_type
	end
	
	while i<no_of_errors
		row_no=@@error_array[i]
		
		######################formatting######################################
		sheet_new.row(0).default_format = format ##changing format
		sheet_new.column(0).default_format=format1
		sheet_new.column(1).default_format=format1		
		#row1=sheet.row(row_no)
		title=sheet_new.row(0)
		title[0]='Row Numbers'
		title[1]='Error Type'
		##############################################################3
		
		
		###############################decoding error type############################33
		if @@error_array_type[i]==1
			error_type='Party\'s Name is invalid'
		end
		if @@error_array_type[i]==2
			error_type='Item name is invalid'
		end
		if @@error_array_type[i]==4
			error_type='Godown name is invalid'
		end
		if @@error_array_type[i]==6
			error_type='Item name and Godown name are invalid'
		end
		if @@error_array_type[i]==3
			error_type='Party\'s Name and Item name are invalid'
		end
		if @@error_array_type[i]==5
			error_type='Party\'s Name and Godown name are invalid'
		end
		if @@error_array_type[i]==7
			error_type='Party\'s Name , Item name, Godown name are all invalid'
		end
		
					#################################################################################
		
		sheet_new[i+2,0]=row_no
		sheet_new[i+2,1]=error_type
		i+=1
	end
		
	book4.write 'errors.xls'

end

###############################################################
def check_error

		for i in 0..(@@n_of_checks-1)
			flag1=0
			flag2=0
			flag3=0
			test=0
			
		
				for j in 0..(@@n_of_master)
						if @@test_name[i]==@@master_name[j]
						flag1=1
						break
						end
				end
		
				for k in 0..(@@n_of_master)
						if @@test_item[i]==@@master_item[k] && @@master_item[k]!=nil
						flag2=1
						break
						end
				end
		
				for l in 0..(@@n_of_master)
						if @@test_godown[i]==@@master_godown[l] && @@master_godown[l]!=nil
						flag3=1
						break
						end
						
				end
		
			if flag1==0||flag2==0||flag3==0
			local_error=@@error_array_index
			@@error_array[local_error]=(i+2)
				if flag1==0
				test+=1
				end
				if flag2==0
				test+=2
				end
				if flag3==0
				test+=4
				end
			@@error_array_type[local_error]=test	
			@@error_array_index+=1
			end
			
		end #end of i loop
			#puts @@error_array
			
	end #end of method


end  ##############end of class
################################################################
##object creation

a=Automation.new
a.parsing_master("master.xls")
a.parsing_checks("udi.xls")
a.check_error

a.error_hilighting





