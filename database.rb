require File.join(File.dirname(__FILE__), 'lib/exasol')
require 'spreadsheet'
require 'yaml'

config = YAML.load_file("config/config.yaml")
@login = config["login"]
@password = config["password"]

#Create result file
result_excel = Spreadsheet::Workbook.new
sheet1 = result_excel.create_worksheet
sheet1.name = 'Result'
sheet1.row(0).concat %w{ID Offer Advertiser Country Approved Rejected %_Of_Rejected_Conversions Potential_Damage Whitelisted? Protocol Multiple_Conversions Statuses_Info}

row_counter = 1

@connection = Exasol.new(@login, @password)
@connection.connect

Spreadsheet.open('result.xls') do |book|
  book.worksheet('Result').each do |row|
    next if row[0] == "Offer_ID"

    query_1 = "select distinct co.code from cms.countries as co join cms.program_regions as pr on co.id = pr.country_id join cms.landing_pages as lp on pr.id = lp.program_region_id
where lp.affiliate_offer_id = '#{row[0]}'"

    @connection.do_query(query_1)
    result_1 = @connection.print_result_array
    puts result_1
    s = String.new
    if result_1.flatten.size > 1
      result_1.flatten.each {|l| s << l + ", "}
      s = s.chop.chop
    else 
      s = result_1.flatten[0]
    end
        if result_1.empty?
          excel_row = sheet1.row(row_counter)
          excel_row[0] = row[0]
          excel_row[1] = row[1]
          excel_row[2] = row[2]
          excel_row[3] = "Offer is not created in AMS"
          excel_row[4] = row[3]
          excel_row[5] = row[4]
          excel_row[6] = row[5]
          excel_row[7] = row[6]
          excel_row[8] = row[7]
          excel_row[9] = row[8]
          excel_row[10] = row[9]
          excel_row[11] = row[10]
        else
          excel_row = sheet1.row(row_counter)
          excel_row[0] = row[0]
          excel_row[1] = row[1]
          excel_row[2] = row[2]
          excel_row[3] = s
          excel_row[4] = row[3]
          excel_row[5] = row[4]
          excel_row[6] = row[5]
          excel_row[7] = row[6]
          excel_row[8] = row[7]
          excel_row[9] = row[8]
          excel_row[10] = row[9]
          excel_row[11] = row[10]
        end

      row_counter += 1

  end

end

@connection.disconnect
result_excel.write 'final_result.xls'
