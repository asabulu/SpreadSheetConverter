require 'rubygems'
require 'spreadsheet' 
require 'dbi'
require 'yaml'
#require 'pp'
#require 'date'

$UseTran = false
#$InsStr = ""

begin
  $Config = YAML.load(File.open("config.yml"))

  $UseTran = $Config["transaction"]

rescue ArgumentError => e
  puts "Could not parse YAML: #{e.message}"
  exit
end

prefix = <<EOB
Set XACT_ABORT ON
GO
BEGIN TRANSACTION
GO

EOB

postfix = <<EOB

COMMIT TRANSACTION
GO
EOB

def readTableSchema (tableName)

  begin
    server = $Config["database"]["server"] #"SAMPLEDB"
    db = $Config["database"]["name"] #"NORTHWIND"
    usr = $Config["database"]["usr"] #"vsuser"
    pwd = $Config["database"]["pwd"] #"vsuser"

    auth = $Config["NTAuthentication"]

    #cn = DBI.connect('dbi:ODBC:DEV_PS','vsuser','vsuser')
    cn = (auth == false)? DBI.connect("DBI:ODBC:Driver={SQL Server};Server=#{server};Database=#{db};Uid=#{usr};Pwd=#{pwd}") : DBI.connect("DBI:ODBC:Driver={SQL Server};Server=#{server};Database=#{db};Trusted_Connection=yes")

    reader = cn.execute("SELECT COLUMN_NAME,IS_NULLABLE,DATA_TYPE,CHARACTER_MAXIMUM_LENGTH,CHARACTER_OCTET_LENGTH,NUMERIC_PRECISION,NUMERIC_PRECISION_RADIX,NUMERIC_SCALE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_CATALOG='#{db}' AND TABLE_NAME = '#{tableName}'")

    hash={}

    reader.each do |row|

      hash[row[0]] = row.to_h

    end  

    cn.disconnect
  rescue DBI::DatabaseError => e
    puts "Read Table Schema error,please check your database connection setting in config.yml"
    exit
  end
  

  return hash

end

def checkValue(nullAble, chkValue , srcValue, srcType)
  mResult = ""
  
  case srcType
  when String,Date
    mResult = "'#{chkValue.gsub(/\'/,"''")}'"
  else
    mResult = chkValue.to_s()
  end
  
  return mResult

end

def convertDBValue(colType, colValue,rowLine)
  mResult = ""
  # p colType
  # p colValue.class
  # 
  #p colType["CHARACTER_OCTET_LENGTH"]
  if (colValue == nil || colValue.to_s() == "" || colValue.to_s().upcase == "NULL")
    if (colType['IS_NULLABLE'] == "NO")
       if (colType["DATA_TYPE"] == "nvarchar" || colType["DATA_TYPE"] == "varchar")
       	  return "''"
       else
      	  puts "\n***** Error: Line::#{rowLine} - Column #{colType['COLUMN_NAME']} is required, but it have null value. *****\n" 
       end
    end
    return "null"
  end  

  case colType["DATA_TYPE"]
  when "nvarchar","varchar","nchar","ntext"
    case colValue
    when Numeric,Float,Fixnum
      mResult = checkValue(colType["IS_NULLABLE"],(colValue.to_i() == colValue) ? colValue.to_i().to_s() : colValue.to_s(), colValue, "")
    when Date
      mResult = checkValue(colType["IS_NULLABLE"], colValue.strftime("%Y/%m/%d %H:%M:%S"), colValue, colValue)
    when String,NilClass
      mResult = checkValue(colType['IS_NULLABLE'], colValue, colValue, "")
    else
      puts "\n***** Error: Unsupport data type in Excel #{colType["COLUMN_NAME"]}:#{colValue.class}...\n"
    end 
    puts "\n***** Warning: The max length of Column[#{colType["COLUMN_NAME"]}] is #{colType["CHARACTER_MAXIMUM_LENGTH"]} < #{mResult}(#{mResult.length}) ... *****\n" if ((colValue.class == String) && (mResult.length() -2 > colType["CHARACTER_MAXIMUM_LENGTH"]))
  when "smallint","int","bigint"
    case colValue
    when Numeric,Float,Fixnum
      mResult = checkValue(colType["IS_NULLABLE"],(colValue.to_i() == colValue) ? colValue.to_i().to_s() : colValue.to_s(), colValue, "")
    when Date,String,NilClass
      mResult = checkValue(colType['IS_NULLABLE'], colValue.to_i(), colValue, 1)
    else
      puts "\n***** Error: Unsupport data type in Excel #{colType["COLUMN_NAME"]}:#{colValue.class}...\n"
    end 
    #puts "error at #{colType['COLUMN_NAME']}"
    puts "\n***** Warning: INT:: The original scale of Column[#{colType['COLUMN_NAME']}] value will be lost (#{colValue} -> #{colValue.to_i()}) ... *****\n" if colValue.to_i() != colValue
  when "decimal"
    case colValue
    when Numeric,Float,Fixnum
      mResult = checkValue(colType["IS_NULLABLE"],(colValue.to_i() == colValue) ? colValue.to_i().to_s() : colValue.to_s(), colValue , 1)
    when Date,String,NilClass
      mResult = checkValue(colType["IS_NULLABLE"],(colValue.to_f().to_s()!=colValue.to_s()) ? "null":colValue, colValue , 1.0)
    else
      puts "\n***** Error: Unsupport data type in Excel #{colType["COLUMN_NAME"]}:#{colValue.class}...\n"
    end 
    puts "\n***** Warning: DECIMAL:: The original scale of Column[#{colType['COLUMN_NAME']}] value will be lost (#{colValue} -> #{colValue.to_f().round(colType['NUMERIC_PRECISION_RADIX'].to_i())}) ... *****\n" if colValue.to_f().round(colType['NUMERIC_PRECISION_RADIX'].to_i()) != colValue
  when "datetime","PSDATE:datetime"
    mResult = (colValue==nil || colValue.to_s()=="") ? "null":checkValue(colType['IS_NULLABLE'], colValue.strftime("%Y/%m/%d %H:%M:%S"),colValue, colValue)
  else
    puts "\n***** Error: Unsupport data type #{colType['COLUMN_NAME']}-#{colType['DATA_TYPE']} ...\n"
  end  

  return mResult

end  

#tstart = Time.now

# if ARGV[0]==nil then
   # puts "Please input Excel file name (Must be .xls type)..."
   # exit
# end 

# ARGV.each do |argvStr|
  # AssignPara(argvStr.to_s().downcase)
# end

#inputName = ARGV[0].dup

puts <<EOB
SpreadSheetConverter ver 1.0.0 by Asa & Jonny
EOB
    sleep 2
	
    inputName = $Config["input"]

begin
  
	book = Spreadsheet.open (inputName.index('.xls') == nil ? inputName << ".xls" : inputName)

rescue StandardError => e
  
	puts "Could not load input file: #{e.message} \n"
    
	puts "Please check the setting in config file and make sure data file is not opened by other application."
	
	exit
	
end    
#p $Config

book.worksheets.each do |activeSheet|
  
  puts "Process for #{ activeSheet.name}..."

  columnInfo = readTableSchema(activeSheet.name)

    File.open("#{activeSheet.name}.sql", 'w') do |f|

      f.write(prefix) if ($UseTran == true)
      
      column_list = activeSheet.row(0).join(',')
      sheet_name = activeSheet.name

      activeSheet.each_with_index 1 do |row, index|
        len=row.length     
        doc=""

        insStr = "INSERT INTO " << sheet_name << "(" << column_list << ") VALUES(" 
        #insStr = "INSERT INTO #{activeSheet.name} (#{column_list}) VALUES("

        (0..len-1).each do |i|
          col=row[i]

          doc << convertDBValue(columnInfo[activeSheet.row(0)[i]], col , index)    
          doc << (",") if i != len-1
        end
        
        f.write("#{insStr << doc});\n")

      end 

      f.write(postfix) if ($UseTran == true)

    end 

end   

#tend = Time.now 

#print "time consume #{(tend - tstart)} "

puts
puts "processing completed."
