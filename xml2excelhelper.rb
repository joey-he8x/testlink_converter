require 'rubygems'
require 'spreadsheet'
require 'nokogiri'
require 'ruby-debug'

class XML2ExcelHelper
  attr_accessor :book,:ws
  TITLE_ROW = 1
  def initialize
  end

  private
  def create_excel_header
    @book = Spreadsheet::Workbook.new 
    @ws = book.create_worksheet
    @ws.insert_row(ws.row_count,["Test Suite Level","","","","","","TC Detail","","","","","","","CUSTOM FIELDS"])
    @ws.insert_row(ws.row_count,["1","2","3","4","","TC ID","NAME","SUMMARY","PRECONDITIONS","STEPS","EXPECTED RESULT","AUTOMATED","KEYWORDS"])
    @ws.merge_cells 0,0,0,4
    @ws.merge_cells 0,6,0,12
  end

  def add_custom_fields  doc
    cf = doc.xpath(".//custom_field/name/text()").map do |name|
      name.text
    end
    cf.uniq!
    @ws.merge_cells 0,13,0,12+cf.size
    cf.each do |name|
      row = @ws.row TITLE_ROW
      row << name
    end
  end

  def write_excel filename
    @book.write filename
  end

  def write_body  xmlfile
    doc = Nokogiri::XML.parse(xmlfile,nil,nil, Nokogiri::XML::ParseOptions::NOBLANKS)
    add_custom_fields doc
    root = doc.child
    if root.attribute("name").nil? or root.attribute("name").value.empty?
      level=0
    else
      level=1
    end
    write_suite root,level
  end

  def write_suite suite_node,level
    if level > 0
      rowid = get_nextrow
      suite_name = suite_node.attribute("name").value
      @ws[rowid,level-1] = suite_name
      @ws[rowid,4] = suite_name
    end
    #处理所有的testcase
    suite_node.xpath("./testcase").each do |tc|
      write_case tc
    end
    #递归处理testsuite
    suite_node.xpath("./testsuite").each do |ts|
      write_suite ts,level+1
    end
 
  end
  def write_case tc
    rowid=get_nextrow
    name = tc.attribute("name").value
    write_cell rowid,"NAME",name
    externalid = tc.xpath("./externalid/text()").text.strip
    write_cell rowid,"TC ID",externalid
    summary = tc.xpath("./summary/text()").text.strip
    write_cell rowid,"SUMMARY",summary
    preconditions = tc.xpath("./preconditions/text()").text.strip
    write_cell rowid,"PRECONDITIONS",preconditions
    execution_type = tc.xpath("./execution_type/text()").text.strip
    if execution_type == "1" then
      write_cell rowid,"AUTOMATED","N"
    else 
      write_cell rowid,"AUTOMATED","Y" if execution_type == "2"
    end
    step_actions = []
    step_results = []
    tc.xpath("./steps/step").each do |step|
      step_actions << step.xpath("./actions/text()").text.strip
      step_results << step.xpath("./expectedresults/text()").text.strip
    end
    write_cell rowid,"STEPS",step_actions.join("\n")
    write_cell rowid,"EXPECTED RESULT",step_results.join("\n")
    keywords=[]
    tc.xpath("./keywords/keyword").each do |keyword|
      keywords << keyword.attribute("name").value
    end
    write_cell rowid,"KEYWORDS",keywords.join(",")
    #custom_fields = {}
    tc.xpath("./custom_fields/custom_field").each do |cfield|
      #custom_fields[cfield.xpath("./name/text()").text] = cfield.xpath("./value/text()").text.strip
      write_cell rowid,cfield.xpath("./name/text()").text,cfield.xpath("./value/text()").text.strip
    end
  end

  def write_cell rowid,column,value
    @ws.column_count.times do |i|
      if @ws[TITLE_ROW,i] == column
        @ws[rowid,i] = value
      end
    end
  end

  def get_nextrow
    @ws.row_count
  end

  public
  def convert infile,outfile
    create_excel_header
    write_body infile
    write_excel outfile
  end
end


if __FILE__ == $0
  ex=XML2ExcelHelper.new
  f=File.open('data/all.xml')
  ex.convert f,"ttt.xls"
end
