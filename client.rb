require File.join(File.dirname(__FILE__), 'lib/requester')
require 'spreadsheet'
require 'yaml'
require 'yajl'

#load hasoffers config
config = YAML.load_file("config/config.yaml")
network_id = config["network_id"]
network_token = config["network_token"]
url = config["url"]

#Create result file
result = Spreadsheet::Workbook.new
sheet1 = result.create_worksheet
sheet1.name = 'Result'
sheet1.row(0).concat %w{Offer_ID Offer_name Advertiser Approved Rejected %_of_rejected_conversions potential_damage offer_whitelisted? offer_protocol allow_multiple_conversions statuses_info}

row_counter = 1

case ARGV[0]
when "yesterday"
  date = (Time.now - 86400).strftime('%Y-%m-%d')
when "today"
  date = Time.now.strftime("%Y-%m-%d")
else
  STDOUT.puts <<-EOF
  Please provide command name

  Usage:
    ruby client.rb yesterday
    ruby client.rb today
  EOF
end

#Call HasOffers API for getting information

response = Requester.make_request(
  url,
  {
  "NetworkId" => network_id,
  "NetworkToken" => network_token,
  "Target" => "Report",
  "Method" => "getConversions",
  "fields[0]" => "Offer.id",
  "fields[1]" => "Offer.name",
  "fields[2]" => "Stat.count_approved",
  "fields[3]" => "Stat.net_payout",
  "fields[4]" => "Stat.count_rejected",
  "fields[5]" => "Stat.rejected_rate",
  "groups[0]" => "Offer.name",
  "filters[Stat.date][conditional]" => "BETWEEN",
  "filters[Stat.date][values][0]" => date,
  "filters[Stat.date][values][1]" => date,
  "sort[Stat.count_rejected]" => "desc",
  "limit" => "50",
  "page" => "1",
  "totals" => "1",
  "data_start" => date,
  "data_end" => date,
  "hour_offset" => "0"
  },
  :get
)

#Parse JSON data

json = StringIO.new(response)
parser = Yajl::Parser.new
hash = parser.parse(json)

hash["response"]["data"]["data"].each do |offer|
  
  offer_info = Requester.make_request(
    url,
    {
    "NetworkId" => network_id,
    "NetworkToken" => network_token,
    "Target" => "Offer",
    "Method" => "findById",
    "id" => offer["Offer"]["id"]
    },
    :get
  )

  json = StringIO.new(offer_info)
  parser = Yajl::Parser.new
  offer_hash = parser.parse(json)

  oh = offer_hash["response"]["data"]["Offer"]

  advertiser_info = Requester.make_request(
    url,
    {
    "NetworkId" => network_id,
    "NetworkToken" => network_token,
    "Target" => "Advertiser",
    "Method" => "findById",
    "id" => oh["advertiser_id"]
    },
    :get
  )

  json = StringIO.new(advertiser_info)
  parser = Yajl::Parser.new
  advertiser_hash = parser.parse(json)

  ah = advertiser_hash["response"]["data"]["Advertiser"]

  status_info = Requester.make_request(
    url,
    {
    "NetworkId" => network_id,
    "NetworkToken" => network_token,
    "Target" => "Report",
    "Method" => "getConversions",
    "fields[0]" => "Stat.status_code",
    "filters[Stat.status][conditional]" => "EQUAL_TO",
    "filters[Stat.status][values][0]" => "rejected",
    "filters[Offer.id][conditional]" => "EQUAL_TO",
    "filters[Offer.id][values][0]" => offer["Offer"]["id"],
    "filters[Stat.date][conditional]" => "BETWEEN",
    "filters[Stat.date][values][0]" => date,
    "filters[Stat.date][values][1]" => date,
    "limit" => "1000",
    "page" => "1",
    "totals" => "1"
    },
    :get
  )

  json = StringIO.new(status_info)
  parser = Yajl::Parser.new
  status_hash = parser.parse(json)

  #statuses variables
  kp = dc = nru = sinw = wtp = apld = dcce = dcbui = et = at = cae = cct = sct = spt = rt = adj = mcce = dpbe = mpbe = drbe = mrbe = amcee = adpbe = ampbe = adrbe = amrbe = cssir = uk = 0

  status_hash["response"]["data"]["data"].each do |status|
    code = status["Stat"]["status_code"]

    case code
    when '11'
      kp += 1
    when '12'
      dc += 1
    when '13'
      nru += 1
    when '14'
      sinw += 1
    when '15'
      wtp += 1
    when '16'
      apld +=1
    when '17'
      dcce += 1
    when '18'
      dcbui +=1
    when '21'
      et += 1
    when '22'
      at += 1
    when '31'
      cae += 1
    when '41'
      cct += 1
    when '42'
      sct += 1
    when '43'
      spt += 1
    when '51'
      rt += 1
    when '52'
      adj += 1
    when '61'
      mcce += 1
    when '62'
      dpbe += 1
    when '63'
      mpbe += 1
    when '64'
      drbe += 1
    when '65'
      mrbe += 1
    when '81'
      amcce += 1
    when '82'
      adpbe += 1
    when '83'
      ampbe += 1
    when '84'
      adrbe += 1
    when '85'
      amrbe += 1
    when '99'
      cssir += 1
    else
      uk += 1
      puts "Unknown statu: #{code}"
    end

  end

  status_result = Hash.new
  
  status_result = { "Known Proxy" => kp, "Duplicate Conversion by Transaction ID" => dc, "No Referral URL" => nru, "Server IP not Whitelisted" => sinw, "Wrong Tracking Protocol" => wtp, 
                    "Affiliate Pixel Loop Detected" => apld, "Daily Conversion Cap Exceeded" => dcce, "Duplicate Conversion by Unique ID" => dcbui, "Employee Test" => et, 
                    "Affiliate Test" => at, "Conversion approval enabled" => cae, "Client cookie tracking" => cct, "Server cookie tracking" => sct, "Server postback tracking" => spt, 
                    "RingRevenue tracking" => rt, "Adjustment" => adj, "Monthly Conversion Cap Exceeded" => mcce, "Daily Payout Budget Exceeded" => dpbe, "Monthly Payout Budget Exceeded" => mpbe, 
                    "Daily Revenue Budget Exceeded" => drbe, "Monthly Revenue Budget Exceeded" => mrbe, "Affiliate Monthly Conversion Cap Exceeded" => amcee, 
                    "Affiliate Daily Payout Budget Exceeded" => adpbe, "Affiliate Monthly Payout Budget Exceeded" => ampbe, "Affiliate Daily Revenue Budget Exceeded" => adrbe, 
                    "Affiliate Monthly Revenue Budget Exceeded" => amrbe, "Conversion Status set in Request" => cssir, "Unknown Status" => uk 
                  }

  status_result.delete_if {|key, value| value == 0}
  status_res = String.new
  status_result.each {|key, value| status_res << "#{key}: #{value}, "}

  case oh["protocol"]
  when "http_img"
    protocol = "Image Pixel"
  when "https_img"
    protocol = "Secure Image Pixel"
  when "server"
    protocol = "Server to Server"
  when "http"
    protocol = "iFrame Pixel"
  when "https"
    protocol = "Secure iFrame Pixeli"
  end

  potential_damage = (oh["default_payout"].to_f * offer["Stat"]["count_rejected"].to_f).round(2)
  if oh["currency"].nil? or oh["currency"].empty? or oh["currency"] == ""
    currency = "EUR"
  else
    currency = oh["currency"]
  end

  puts oh["currency"]

  damage = potential_damage.to_s + ' ' + currency

  sheet1.row(row_counter).push offer["Offer"]["id"], offer["Offer"]["name"], ah["company"], offer["Stat"]["count_approved"], offer["Stat"]["count_rejected"], (offer["Stat"]["rejected_rate"].to_f).round(2), damage, oh["enable_offer_whitelist"] == "0" ? "No" : "Yes", protocol, oh["allow_multiple_conversions"] == "0" ? "No" : "Yes", status_res.chop.chop

  row_counter += 1
end

result.write 'result.xls'
