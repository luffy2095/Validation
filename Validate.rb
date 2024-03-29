#!/usr/bin/ruby
require 'spreadsheet'
require 'rubyXL'
require 'amatch'
include Amatch
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
@@limit=2


#####################################################definitions##################################

def parsing_master(f)

#################################parsing xls first##############################################3

	if f =~ /\b.xls$\b/                                  # check if file is xls format

	
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
			
	@@n_of_master=no_of_master_parsed
	#puts no_of_master_parsed
	end			
######################parsing of xls comp###############################

####################parsing xlsx########################################
	if f =~ /\b.xlsx$\b/
	
		book1 = RubyXL::Parser.parse(f)
		sheet1=book1[0]
		act_sheet=sheet1.extract_data
		no_of_master_parsed=0
		i=0
		act_sheet.each  do |row|
			if i==0
				i+=1
				next
			end
			@@master_name[no_of_master_parsed]=row[0].to_s
			@@master_item[no_of_master_parsed]=row[1].to_s
			@@master_godown[no_of_master_parsed]=row[2].to_s
			no_of_master_parsed+=1
			
		end
		@@n_of_master=no_of_master_parsed
		
	end		
################### xlsx comp#######################################333
	
#		puts no_of_master_parsed
#		puts @@master_item	
		
end		
#####################3end of parsing_master#######################################33
  

#########################################################

def parsing_checks(f)
	
####################################3parsing xls first##########################################	
	if f =~ /\b.xls$\b/
	
		no_of_test_parsed=0
		book2 = Spreadsheet.open(f,'rb')
		sheet = book2.worksheet 0 ##parse each worksheet
	
		sheet.each 1 do |row|
		
			if row[4]==nil 
				break
			end
			@@test_name[no_of_test_parsed]=row[4]			########check for errors
			@@test_item[no_of_test_parsed]=row[5]
			@@test_godown[no_of_test_parsed]=row[20]
			no_of_test_parsed+=1
		end		###end of row parsing
		
		@@n_of_checks=no_of_test_parsed
		
		end  #############end of xls parsing_checks
		
##############parsing xlsx now###################################################################3

	if f =~ /\b.xlsx$\b/
		no_of_test_parsed=0
		book2= RubyXL::Parser.parse(f)
		sheet=book2[0]
		act_sheet=sheet.extract_data
		
		i=0
		act_sheet.each do |row|
			if i==0
				i+=1
				next
			end
			if row[4]==nil
				break
			end
			@@test_name[no_of_test_parsed]=row[4]					########check for errors
			@@test_item[no_of_test_parsed]=row[5]
			@@test_godown[no_of_test_parsed]=row[20]
			no_of_test_parsed+=1
		end
		
		@@n_of_checks=no_of_test_parsed
	end
#####################################end of xlsx parsing#########################################
				
end		#####end of parsing_checks



############################################################

def error_hilighting(f)

	if f =~ /\b.xls\b/
		no_of_errors=(@@error_array_index)
	
		#book3=Spreadsheet.open(f,'rb') ###opening for editing
		#sheet = book3.worksheet 0
	
	
		book4=Spreadsheet::Workbook.new   ##creating new sheet
		sheet_new=book4.create_worksheet
	
		sheet_new.column(0).width=20  #row no
		sheet_new.column(1).width=50	#name
		sheet_new.column(2).width=30	#item
		sheet_new.column(3).width=30	#godown
		sheet_new.column(4).width=60	#error
		sheet_new.column(5).width=40	#suggestion
		
	
	
		format = Spreadsheet::Format.new    :color => :blue,
							:weight => :bold,
							:size => 20,
							:align => :center
						
		format1 = Spreadsheet::Format.new   :align => :center,
							 :size => 10
							
		i=0
		if no_of_errors==0
			sheet_new.row(0).default_format = format ##changing format
			sheet_new.column(0).default_format=format1
			sheet_new.column(1).default_format=format1
			sheet_new.column(2).default_format=format1	
			sheet_new.column(3).default_format=format1
			sheet_new.column(4).default_format=format1
			sheet_new.column(5).default_format=format1				
			#row1=sheet.row(row_no)
			title=sheet_new.row(0)
			title[0]='Row No'
			title[1]='Name'
			title[2]='Item'
			title[3]='Godown'
			title[4]='Error'
			title[5]='Suggested Correction'
			error_type='There are no errors in the UDI file.All entries are valid'
			sheet_new[i+2,4]=error_type
		end
	
		while i<no_of_errors
			row_no=@@error_array[i]
		
			######################formatting######################################
			sheet_new.row(0).default_format = format ##changing format
			sheet_new.column(0).default_format=format1
			sheet_new.column(1).default_format=format1
			sheet_new.column(2).default_format=format1	
			sheet_new.column(3).default_format=format1
			sheet_new.column(4).default_format=format1
			sheet_new.column(5).default_format=format1				
			#row1=sheet.row(row_no)
			title=sheet_new.row(0)
			title[0]='Row No'
			title[1]='Name'
			title[2]=' Item'
			title[3]='Godown'
			title[4]='Error'
			title[5]='Suggested Correction'
			##############################################################3
		
		
			###############################decoding error type############################33
		if @@error_array_type[i]==1
			error_type='Party\'s Name is invalid'
			suggestion=correct_error(1,i)
		end
		if @@error_array_type[i]==2
			error_type='Item name is invalid'
			suggestion=correct_error(2,i)
		end
		if @@error_array_type[i]==4
			error_type='Godown name is invalid'
			suggestion=correct_error(4,i)
		end
		if @@error_array_type[i]==6
			error_type='Item name and Godown name are invalid'
			suggestion=correct_error(6,i)
		end
		if @@error_array_type[i]==3
			error_type='Party\'s Name and Item name are invalid'
			suggestion=correct_error(3,i)
		end
		if @@error_array_type[i]==5
			error_type='Party\'s Name and Godown name are invalid'
			suggestion=correct_error(5,i)
		end
		if @@error_array_type[i]==7
			error_type='Party\'s Name , Item name, Godown name are all invalid'
			suggestion=correct_error(7,i)
		end
		#############################
		if @@error_array_type[i]==10
			error_type='Item in wrong godown'
			suggestion=correct_error(10,i)
		end
		if @@error_array_type[i]==11
			error_type='Party\'s Name is invalid; item in wrong godown'
			suggestion=correct_error(11,i)
		end
		if @@error_array_type[i]==12
			error_type='Item name is invalid; item in wrong godown'
			suggestion=correct_error(12,i)
		end
		if @@error_array_type[i]==14
			error_type='Godown name is invalid; item in wrong godown'
			suggestion=correct_error(14,i)
		end
		if @@error_array_type[i]==16
			error_type='Item name and Godown name are invalid; item in wrong godown'
			suggestion=correct_error(16,i)
		end
		if @@error_array_type[i]==13
			error_type='Party\'s Name and Item name are invalid; item in wrong godown'
			suggestion=correct_error(13,i)
		end
		if @@error_array_type[i]==15
			error_type='Party\'s Name and Godown name are invalid; item in wrong godown'
			suggestion=correct_error(15,i)
		end
		if @@error_array_type[i]==17
			error_type='Party\'s Name , Item name, Godown name are all invalid; item in wrong godown'
			suggestion=correct_error(17,i)
		end
						#################################################################################
		
			sheet_new[i+2,0]=row_no
			sheet_new[i+2,1]=@@test_name[row_no-2]
			sheet_new[i+2,2]=@@test_item[row_no-2]
			sheet_new[i+2,3]=@@test_godown[row_no-2]
			sheet_new[i+2,4]=error_type
			if suggestion!=nil
				sheet_new[i+2,5]=suggestion
			end
			i+=1
		end
		
		book4.write 'errors_xls.xls'

	end
	################################xlsx parsing############################3
	
	
##		file=File.open(f,"r+")
#			puts "i am here"
#		book_edit = RubyXL::Parser.parse(f) 
#			puts "but not here"
#			book_edit.each do |sheet_edit|
#	
#			no_of_errors=@@error_array_index
#			i=0
#			while i<no_of_errors
#				row_no=@@error_array[i]
#				sheet_edit.change_row_bold(row_no-1,true)
#				sheet_edit.change_row_italics(row_no-1,true)
#				print " changed "
#				puts row_no
#	#			sheet_edit.change_row_bold(row_no,true)
#				i+=1
#			end
#			book_edit.write(f)
#		
#		
#			end
#	
#	
#		#puts "file deleted"
##		file.close

	if f=~ /\b.xlsx\b/
		no_of_errors=@@error_array_index
		book_new=RubyXL::Workbook.new
		
		sheet_new=book_new.worksheets[0]
		
		sheet_new.change_column_width(0,15)	#row
		sheet_new.change_column_width(1,45)	#name
		sheet_new.change_column_width(2,25)	#item
		sheet_new.change_column_width(3,25)	#godown
		sheet_new.change_column_width(4,60)	#error
		sheet_new.change_column_width(5,35)	#suggestion
		sheet_new.change_column_horizontal_alignment(0,'center')
		sheet_new.change_column_horizontal_alignment(1,'center')
		sheet_new.change_column_horizontal_alignment(2,'center')
		sheet_new.change_column_horizontal_alignment(3,'center')
		sheet_new.change_column_horizontal_alignment(4,'center')
		sheet_new.change_column_horizontal_alignment(5,'center')
		sheet_new.change_column_font_size(0,10)
		sheet_new.change_column_font_size(1,10)
		sheet_new.change_column_font_size(2,10)
		sheet_new.change_column_font_size(3,10)
		sheet_new.change_column_font_size(4,10)
		sheet_new.change_column_font_size(5,10)
		sheet_new.change_row_bold(0,true)
		sheet_new.change_row_font_size(0,20)
		sheet_new.change_row_font_name(0,'Arial')
		sheet_new.change_row_horizontal_alignment(0,'center')
		sheet_new.change_row_font_color(0,'0000ff')
		
		i=0
		if no_of_errors==0
			
			#title=sheet_new[0]
			#title[0]='Row Numbers'
			#title[1]='Error Type'
			sheet_new.add_cell(0,0,'Row No ')
			sheet_new.add_cell(0,1,' Name ')
			sheet_new.add_cell(0,2,' Item ')
			sheet_new.add_cell(0,3,' Godown ')
			sheet_new.add_cell(0,4,' Errors ')
			sheet_new.add_cell(0,5,' Suggested Correction')
			error_type='There are no errors in the UDI file.All entries are valid'
			sheet_new.add_cell(2,1,error_type)
			#message[1]=error_type
		end
		
		while i<no_of_errors
			row_no=@@error_array[i]
			sheet_new.add_cell(0,0,'Row No ')
			sheet_new.add_cell(0,1,' Name ')
			sheet_new.add_cell(0,2,' Item ')
			sheet_new.add_cell(0,3,' Godown ')
			sheet_new.add_cell(0,4,' Errors ')
			sheet_new.add_cell(0,5,' Suggested Correction')
			######################formatting######################################
#			sheet_new.row(0).default_format = format ##changing format
#			sheet_new.column(0).default_format=format1
#			sheet_new.column(1).default_format=format1		
#			#row1=sheet.row(row_no)
#			title=sheet_new[0]
			#sheet_new[0][0]='Row Numbers'
			#sheet_new[0][1]='Error Type'
			##############################################################3
		
		
			###############################decoding error type############################33
					if @@error_array_type[i]==1
						error_type='Party\'s Name is invalid'
						suggestion=correct_error(1,i)
					end
					if @@error_array_type[i]==2
						error_type='Item name is invalid'
						suggestion=correct_error(2,i)
					end
					if @@error_array_type[i]==4
						error_type='Godown name is invalid'
						suggestion=correct_error(4,i)
					end
					if @@error_array_type[i]==6
						error_type='Item name and Godown name are invalid'
						suggestion=correct_error(6,i)
					end
					if @@error_array_type[i]==3
						error_type='Party\'s Name and Item name are invalid'
						suggestion=correct_error(3,i)
					end
					if @@error_array_type[i]==5
						error_type='Party\'s Name and Godown name are invalid'
						suggestion=correct_error(5,i)
					end
					if @@error_array_type[i]==7
						error_type='Party\'s Name , Item name, Godown name are all invalid'
						suggestion=correct_error(7,i)
					end
					#############################
					if @@error_array_type[i]==10
						error_type='Item in wrong godown'
						suggestion=correct_error(10,i)
					end
					if @@error_array_type[i]==11
						error_type='Party\'s Name is invalid; item in wrong godown'
						suggestion=correct_error(11,i)
					end
					if @@error_array_type[i]==12
						error_type='Item name is invalid; item in wrong godown'
						suggestion=correct_error(12,i)
					end
					if @@error_array_type[i]==14
						error_type='Godown name is invalid; item in wrong godown'
						suggestion=correct_error(14,i)
					end
					if @@error_array_type[i]==16
						error_type='Item name and Godown name are invalid; item in wrong godown'
						suggestion=correct_error(16,i)
					end
					if @@error_array_type[i]==13
						error_type='Party\'s Name and Item name are invalid; item in wrong godown'
						suggestion=correct_error(13,i)
					end
					if @@error_array_type[i]==15
						error_type='Party\'s Name and Godown name are invalid; item in wrong godown'
						suggestion=correct_error(15,i)
					end
					if @@error_array_type[i]==17
						error_type='Party\'s Name , Item name, Godown name are all invalid; item in wrong godown'
						suggestion=correct_error(17,i)
					end						#################################################################################
		
#			sheet_new[i+2,0]=row_no
#			sheet_new[i+2,1]=error_type

			sheet_new.add_cell(i+2,0,row_no)
			sheet_new.add_cell(i+2,1,@@test_name[row_no-2])
			sheet_new.add_cell(i+2,2,@@test_item[row_no-2])
			sheet_new.add_cell(i+2,3,@@test_godown[row_no-2])
			sheet_new.add_cell(i+2,4,error_type)
			#sheet_new.add_cell(i+2,5,error_type)
			if suggestion!=nil
				sheet_new.add_cell(i+2,5,suggestion)
			end
			i+=1
		end			######end of while
		
		book_new.write 'errors_xlsx.xlsx'

	end				###end of if(to decide betn xls and xlsx)
		
		
		
	

end				###end of method

###############################################################


def check_godown(godown,item)
	godown_error=0
	if godown=='Media'
		case item
			when 'Books','Movies & Music','Posters'
				godown_error=1
			else
				godown_error=0
		end
	end
	
	if godown=='Electronics'
		case item
			when 'Cameras','Camera_Accessories','Computers_Accessories','Gaming_Consoles','Home Appliances','Home Entertainment','Mobiles','Mobile_Accessories','Portable_Electronics','Auto_Accessories'
				godown_error=1
			else
				godown_error=0
		end
	end

	if godown=='Lifestyle'
		case item
			when 'Apparel','Sodexo Gift Voucher','Toys_Games','Shoppers Stop Gift Voucher','Apparel_Accessories','Baby','Bags','Beauty','Eyewear','Life Style Gift Voucher','Grocery','Home_Furnishing','Health Equipments','Kitchenware','Watches','Footwear','Office Suppliers'
				godown_error=1
			else
				godown_error=0
		end
	end
return godown_error
end
#################################
def check_error

		for i in 0..(@@n_of_checks-1)
			flag1=0
			flag2=0
			flag3=0
			flag4=0
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
						if @@test_godown[i]==@@master_godown[l] || @@test_godown[i]==nil
						flag3=1
						break
						end
						
				end
			flag4=check_godown(@@test_godown[i].to_s,@@test_item[i].to_s)
			
			
			if flag1==0||flag2==0||flag3==0||flag4==0
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
				if flag4==0
				test+=10
				end
			@@error_array_type[local_error]=test	
			@@error_array_index+=1
			end
			
		end #end of i loop
			#puts @@error_array
			
end #end of method
############################################################################################################################

def correct_error(abc,index)#######################################################################################3 CHANGE
row_no=@@error_array[index]-2
suggestion=""
		if abc==1 || abc==11
			m=Levenshtein.new(@@test_name[row_no].to_s)
				for j in 0..(@@n_of_master)
				x=m.match(@@master_name[j].to_s)
					if x<=@@limit
						suggestion<<@@master_name[j].to_s
						break
					end
				end
			
			if abc==11
				case @@test_item[row_no]
					when 'Cameras','Camera_Accessories','Computers_Accessories','Gaming_Consoles','Home Appliances','Home Entertainment','Mobiles','Mobile_Accessories','Portable_Electronics','Auto_Accessories'
						suggestion<<" , Electronics as godown"
					when 'Apparel','Toys_Games','Sodexo Gift Voucher','Shoppers Stop Gift Voucher','Apparel_Accessories','Baby','Bags','Beauty','Eyewear','Life Style Gift Voucher','Grocery','Home_Furnishing','Health Equipments','Kitchenware','Watches','Footwear','Office Suppliers'
						suggestion<<" , Lifestyle as godown"
					when 'Books','Movies & Music','Posters'
						suggestion<<" , Media as godown"
				end
			end
		end
		###################
		if abc==10
				case @@test_item[row_no]
					when 'Cameras','Camera_Accessories','Computers_Accessories','Gaming_Consoles','Home Appliances','Home Entertainment','Mobiles','Mobile_Accessories','Portable_Electronics','Auto_Accessories'
						suggestion<<"Electronics as godown"
					when 'Apparel','Toys_Games','Sodexo Gift Voucher','Shoppers Stop Gift Voucher','Apparel_Accessories','Baby','Bags','Beauty','Eyewear','Life Style Gift Voucher','Grocery','Home_Furnishing','Health Equipments','Kitchenware','Watches','Footwear','Office Suppliers'
						suggestion<<"Lifestyle as godown"
					when 'Books','Movies & Music','Posters'
						suggestion<<"Media as godown"
				end
		end
		###################
		if abc==2 || abc==12
			m=Levenshtein.new(@@test_item[row_no].to_s)
				for j in 0..(@@n_of_master)
				x=m.match(@@master_item[j].to_s)
					if x<=@@limit
						suggestion<<@@master_item[j].to_s
						break
					end
				end			
		end
		####################
		if abc==4 || abc==14
			m=Levenshtein.new(@@test_godown[row_no].to_s)
				for j in 0..(@@n_of_master)
				x=m.match(@@master_godown[j].to_s)
					if x<=@@limit
						suggestion<<@@master_godown[j].to_s
						break
					end
				end
		end
		####################
		if abc==6 || abc==16
			m=Levenshtein.new(@@test_item[row_no].to_s)
				for j in 0..(@@n_of_master)
				x=m.match(@@master_item[j].to_s)
					if x<=@@limit
						suggestion<<@@master_item[j].to_s
						break
					end
				end
				
			n=Levenshtein.new(@@test_godown[row_no].to_s)
				for j in 0..(@@n_of_master)
				y=n.match(@@master_godown[j].to_s)
					if y<=@@limit
						suggestion<<"  , "
						suggestion<<@@master_godown[j].to_s
						break
					end
				end
		end
		####################
		if abc==3 || abc==13
			n=Levenshtein.new(@@test_name[row_no].to_s)
				for j in 0..(@@n_of_master)
				y=n.match(@@master_name[j].to_s)
					if y<=@@limit
						suggestion<<@@master_name[j].to_s
						break
					end
				end
				
			m=Levenshtein.new(@@test_item[row_no].to_s)
				for j in 0..(@@n_of_master)
				x=m.match(@@master_item[j].to_s)
					if x<=@@limit
						suggestion<<"  , "
						suggestion<<@@master_item[j].to_s
						break
					end
				end
		end
		#####################
		if abc==5 || abc==15
			n=Levenshtein.new(@@test_name[row_no].to_s)
				for j in 0..(@@n_of_master)
				y=n.match(@@master_name[j].to_s)
					if y<=@@limit
						suggestion<<@@master_name[j].to_s
						break
					end
				end
				
			m=Levenshtein.new(@@test_godown[row_no].to_s)
				for j in 0..(@@n_of_master)
				x=m.match(@@master_godown[j].to_s)
					if x<=@@limit
						suggestion<<"  , "
						suggestion<<@@master_godown[j].to_s
						break
					end
				end
		end
		######################
		if abc==7 || abc==17
			n=Levenshtein.new(@@test_name[row_no].to_s)
				for j in 0..(@@n_of_master)
				y=n.match(@@master_name[j].to_s)
					if y<=@@limit
						suggestion<<@@master_name[j].to_s
						break
					end
				end
				
			o=Levenshtein.new(@@test_item[row_no].to_s)
				for j in 0..(@@n_of_master)
				z=o.match(@@master_item[j].to_s)
					if z<=@@limit
						suggestion<<"  , "
						suggestion<<@@master_item[j].to_s
						break
					end
				end
				
			m=Levenshtein.new(@@test_godown[row_no].to_s)
				for j in 0..(@@n_of_master)
				x=m.match(@@master_godown[j].to_s)
					if x<=@@limit
						suggestion<<"  , "
						suggestion<<@@master_godown[j].to_s
						break
					end
				end
		end
		#############################
		return suggestion
end########### end of method



end  ##############end of class
################################################################
##object creation
##################user input for master ############



#puts "Location of Master File \n 1. Present working directory \n 2. Somewhere else "
#option = gets.chomp
#system "clear"
#begin
#	if option=="1"
#		puts "Please enter master file name with extention.(Case sensitive)"
#		path_master=gets.chomp
#		if(!File.file?(path_master))
#			
#			system "clear"
#			raise
#		end
#	end
#rescue
#puts "please check file name and try again \n"
#retry
#end

#begin
#	if option=="2"
#		puts "Please enter master file name with extention and complete path .(Case sensitive)"
#		path_master=gets.chomp
#		if(!File.file?(path_master))
#			
#			system "clear"
#			raise
#		end
#	end
#rescue
#puts "please check file name and try again \n"
#retry
#end



###########################################################################################
#system "clear"
#######################input for udi #######################################################
puts "Location of UDI File \n 1. Present working directory \n 2. Somewhere else "
option = gets.chomp
system "clear"
begin
	if option=="1"
		puts "Please enter UDI file name with extention.(Case sensitive)"
		path_udi=gets.chomp
		if(!File.file?(path_udi))
			
			#system "clear"
			raise
		end
	end
rescue
puts "please check file name and try again \n"
retry
end


begin
	if option=="2"
		puts "Please enter UDI file name with extention and complete path .(Case sensitive)"
		path_udi=gets.chomp
		if(!File.file?(path_udi))
			
			#system "clear"
			raise
		end
	end
rescue
puts "please check file name and try again \n"
retry
end


##################################################
a=Automation.new
a.parsing_master("master_new.xls")
puts "master parsed"
a.parsing_checks(path_udi)
puts"udi parsed"
a.check_error
puts "cross referencing"
a.error_hilighting(path_udi)
puts "check error files in C:\\Validation folder"




