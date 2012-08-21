require 'rubygems'
require 'sinatra'
require 'xml2excelhelper'
require 'excel2xmlhelper'
$TMPDIR="tmp"

configure do
  mime_type :excel,'application/vnd.ms-excel'
  mime_type :xml,"application/xml"
  mime_type :data,"multipart/form-data"
end

before do
  #content_type :html, 'charset' =>'gb2312'
end

get '/' do
  haml :index
end

post '/convert_xml.xls' do
  unless params[:file] &&
    (tmpfile = params[:file][:tempfile]) &&
    (name = params[:file][:filename])
    @error = "No file selected"
    redirect to("/") 
  end
  ex=XML2ExcelHelper.new
  post_fix=Time.now.to_i
  fname = "#{$TMPDIR}/#{name}#{post_fix}.xls"
  ex.convert tmpfile,fname
  send_file fname,:type => :excel
end

post '/convert_excel.xml' do
  unless params[:file] &&
    (tmpfile = params[:file][:tempfile]) &&
    (name = params[:file][:filename])
    @error = "No file selected"
    redirect to("/") 
  end
  ex=Excel2XMLHelper.new tmpfile
  content_type :data
  return ex.convert
end
