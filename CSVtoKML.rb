require 'geocoder'  # gem install geocoder
require 'csv'       # built-in
require 'kml'       # gem install schleyfox-ruby_kml --source http://gems.github.com
require 'optparse'  # gem install OptionParser
require 'pp'

options = {}

# set defaults - TODO: Make this dynamic
options[:address_column_name] = "Address"
options[:lat_column_name] = "lat"
options[:lng_column_name] = "lng"
options[:url_column_name] = "href"
options[:base_url] = ""
options[:name_column_name] = 0
options[:description_column_name] = "description"
options[:separating_character] = "\t" # "\t"
options[:in_file_name] = "Default.tsv"
options[:out_file_name] = options[:in_file_name]+"-out.txt"
options[:kml_file_name] = options[:in_file_name]+".kml"
options[:map_name] = 'Map'
options[:engine] = :nominatim
options[:text_to_add] = ""

option_parser = OptionParser.new do |opts|
  executable_name = File.split($0)[1]
  opts.banner = "Add column of latitudes and longtitudes for addresses in a spreadsheet and generate a corresponding KML file.
  Usage: #{executable_name} [options]
  "
  # Create a switch
  opts.on("-a ADDRESS_COLUMN_NAME","--address-column-name ADDRESS_COLUMN_NAME", "Name of the column containing addresses.\n Default: #{options[:address_column_name]}") do |column|
    options[:address_column_name] = column
  end
  opts.on("-u URL_COLUMN_NAME","--link-column-name UTL_COLUMN_NAME", "Name of the column containing the link/url.\n Default: #{options[:url_column_name]}") do |column|
    options[:url_column_name] = column
  end
  opts.on("-b BASE_URL","--base-url BASE_URL", "Base URL.\n Default: #{options[:url_column_name]}") do |column|
    options[:base_url] = column
  end
  opts.on("-n NAME_COLUMN_NAME","--name-column-name NAME_COLUMN_NAME", "Name of the column containing name.\n Default: #{options[:name_column_name]}") do |column|
    options[:name_column_name] = column
  end
  opts.on("-d DESCRIPTION_COLUMN_NAME","--description_column_name DESCRIPTION_COLUMN_NAME", "Name of the column containing description.\n Default: #{options[:description_column_name]}") do |column|
    options[:name_column_name] = column
  end
  opts.on("-s SEPARATING_CHARACTER", "--separating-character SEPARATING_CHARACTER", "Separating character in the files.\n Default: #{options[:separating_character]=="\t" ? "tab" : options[:separating_character]}") do |char|
    options[:separating_character] = char
  end
  opts.on("-f IN_FILE", "--in-file IN_FILE", "File to process.\n Default: #{options[:in_file_name]}") do |file|
    options[:in_file_name] = file
    fileparts = file.split(".")
    # puts fileparts
    fileparts.length>1 ? fileparts[-2] = fileparts[-2]+"-out" : fileparts[0] = fileparts[0]+"-out"
    options[:out_file_name] = fileparts.join(".")
    options[:kml_file_name] = options[:in_file_name]+".kml"
  end
  opts.on("-m MAP-NAME", "--map-name MAP_NAME", "Name of the map.\n Default: #{options[:map_name]}") do |name|
    options[:map_name] = name
  end
  opts.on("-e ENGINE-NAME", "--engine-name ENGINE_NAME", "Name of the geocoding engine eg google, yahoo.\n Default: #{options[:engine].to_s}") do |engine|
    options[:engine] = engine.to_sym
  end
  opts.on("-t TEXT_TO_ADD_TO_ADDRESS", "--text-to-add TEXT_TO_ADD_TO_ADDRESS", "Text to add to addresses.\n Default: #{options[:text_to_add].to_s}") do |text|
    options[:text_to_add] = text
  end
  opts.on("-h", "--help", "Display help functions.") do 
    puts option_parser.help
    exit 0
  end
end

begin
  option_parser.parse!
rescue OptionParser::InvalidArgument => ex
  STDERR.puts ex.message
  STDERR.puts option_parser
end

unless File.exist?(options[:in_file_name])
  puts "ERROR: No such file: #{options[:in_file_name]}.\n---\n"
  puts option_parser.help
  exit 1
end

# open files
in_file = CSV.open(options[:in_file_name], headers: true, col_sep: options[:separating_character])
out_file = CSV.open(options[:out_file_name], "w+", headers: true, col_sep: options[:separating_character])
kml_file = File.open(options[:kml_file_name], "w+")

# create kml doc
kml = KMLFile.new
folder = KML::Folder.new(:name => options[:map_name])

# set up the Geocoder
Geocoder::Configuration.lookup = options[:engine]
Geocoder::Configuration.timeout = 2 if options[:engine] == :nominatim

# go through each entry in the in_file
address_column_name = options[:address_column_name]
coordinates_defined = false
first = true
number_of_records = 0
geo_cache = {}
in_file.each do |row|
  if (first)
    headers = row.headers
    unless headers.include? address_column_name
      pp headers
      puts "No such column name: #{address_column_name}.\n---\n"
      puts option_parser.help
      exit 1
    end
    options[:name_column_name] ||= headers[0]
    if headers.include?(options[:lat_column_name])&&(headers.include? options[:lng_column_name])
      puts "Coorinates defined."
      coordinates_defined = true
    else
      headers << options[:lat_column_name] << options[:lng_column_name]
    end
    out_file.puts(headers)
    first = false
  end
  # find latitude and longtidure
  lat = row[options[:lat_column_name]]
  lng = row[options[:lng_column_name]]
  unless coordinates_defined
    geo = geo_cache[row[address_column_name]]
    if geo.nil?
      geo = Geocoder.search(row[address_column_name] + options[:text_to_add])      
      geo_cache[row[address_column_name]] = geo
    end
    unless geo.empty?
        lat = geo[0].coordinates[0]
        lng = geo[0].coordinates[1]
    end
    row << lat << lng
  end
  # write it to the csv-file  
  out_file.puts(row)
  number_of_records+=1

  # add it to the KML document
  folder.features << KML::Placemark.new(
    :name => row[options[:name_column_name]],  # first column is added as a name of the placemark...
    :geometry => KML::Point.new(:coordinates => {:lat => lat, :lng => lng}),
    # :link => options[:base_url] + row[options[:url_column_name]],
    # :description => row[options[:description_column_name]],
  )
end

kml.objects << folder
kml_file.puts kml.render

puts "Latitudes and longtitudes written to: #{options[:out_file_name]}."
puts "KML data written to: #{options[:kml_file_name]}."
puts "Thank you for using Morten's CVStoKML."