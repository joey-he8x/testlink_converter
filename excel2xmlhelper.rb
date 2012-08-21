require 'rubygems'
require 'spreadsheet'
require 'nokogiri'

require 'ruby-debug'

class Excel2XMLHelper
  attr_accessor :doc,:book,:ws

  TITLE_ROW = 1
  def initialize file
    @book = Spreadsheet.open file
    @ws = @book.worksheet(0)
  end

  #private
  def find_child_suites parent_rowid,level
    for i in (parent_rowid+1)..@ws.row_count-1 do
      if !@ws[i,4].nil? and !@ws[i,4].empty? then
        child_level = get_suite_level i
        if child_level == level+1 then
          yield i
        elsif child_level == level then
          return
        end
      end
    end
  end

  def find_child_cases suite_id
    rowid = suite_id + 1
    while (!read_cell(rowid,"NAME").nil?) && (!read_cell(rowid,"NAME").empty?) do
      yield rowid
      rowid +=1
    end
  end

  def get_suite_level rowid
    for i in 0..3
      if !@ws[rowid,i].nil? && !@ws[rowid,i].empty? then
        return i+1
      end
    end
    raise Exception
  end

  def find_root_suites
    for i in (TITLE_ROW+1)..@ws.row_count-1 do
      if (!@ws[i,0].nil?) && (!@ws[i,0].empty?) then
        yield(i)
      end
    end
  end

  def gen_suite suite_id,level
    suite = Nokogiri::XML::Node.new "testsuite",@doc
    suite["name"] = @ws[suite_id,4]
    #处理所有的testcase
    find_child_cases(suite_id) do |rowid|
      suite.add_child(gen_case(rowid))
    end
    #递归处理testsuite
    find_child_suites(suite_id,level) do |rowid|
      suite.add_child(gen_suite(rowid,level+1))
    end
    suite
  end

  def gen_case rowid
    tc = Nokogiri::XML::Node.new "testcase",@doc
    name = read_cell rowid,"NAME"
    tc["name"] = name.to_s
    tc_id = read_cell rowid,"TC ID"
    if !tc_id.nil? and !tc_id.empty? then
      tc.add_child(gen_node("externalid",tc_id))
    end
    summary = read_cell rowid,"SUMMARY"
    tc.add_child gen_node("summary",summary)
    preconditions = read_cell rowid,"PRECONDITIONS"
    tc.add_child gen_node("preconditions",preconditions)
    execution_type = read_cell rowid,"AUTOMATED"
    if execution_type == "N" then
      tc.add_child gen_node("execution_type","1")
    else 
      if execution_type == "Y" then
        tc.add_child gen_node("execution_type","2")
      end
    end
    step_actions = read_cell(rowid,"STEPS")
    if step_actions.nil? then
      step_actions=[]
    else 
      step_actions=step_actions.split
    end
    results = read_cell(rowid,"EXPECTED RESULT")
    if results.nil? then
      results=[]
    else 
      results=results.split
    end
    i = 0
    steps = Nokogiri::XML::Node.new "steps",@doc
    while i < step_actions.size or i < results.size do
      step = Nokogiri::XML::Node.new "step",@doc
      step.add_child gen_node("step_number",(i+1).to_s)
      step.add_child gen_node("actions",step_actions[i])
      step.add_child gen_node("expectedresults",results[i])
      step.add_child gen_node("execution_type","1")
      steps.add_child step
      i+= 1
    end
    tc.add_child steps
    keywords = read_cell(rowid,"KEYWORDS")
    if !keywords.nil? and !keywords.empty? then
      keywords = keywords.split(",")
      keywords_node = Nokogiri::XML::Node.new "keywords",@doc
      keywords.each do |k|
        keyword = Nokogiri::XML::Node.new "keyword",@doc
        keyword["name"] = k
        keyword.add_child gen_node("notes",k)
        keywords_node.add_child keyword
      end
      tc.add_child keywords_node
    end
    custom_fields = Nokogiri::XML::Node.new "custom_fields",@doc
    tc.add_child custom_fields

    
    get_custom_field_range.each do |i|
      name = @ws[TITLE_ROW,i] 
      value = @ws[rowid,i]
      custom_field = Nokogiri::XML::Node.new "custom_field",@doc
      custom_field.add_child gen_node("name",name)
      custom_field.add_child gen_node("value",value)
      custom_fields.add_child custom_field
    end


    tc
  end

  def gen_node name,value
    node = Nokogiri::XML::Node.new name,@doc
    node.add_child @doc.create_cdata(value)
    node
  end

  def read_cell rowid,column
    @ws.column_count.times do |i|
      if @ws[TITLE_ROW,i] == column
        return @ws[rowid,i]
      end
    end
    raise Exception.new "column not found"
  end

  def get_custom_field_range
    hr = @ws.row 0
    tr = @ws.row TITLE_ROW
    start_id = -1
    end_id = 999
    hr.count.times do |i|
      if !hr[i].nil? and start_id != -1
        end_id = i
      end
      if hr[i] == 'CUSTOM FIELDS'
        start_id=i
      end
    end
    end_id = tr.count - 1
    start_id..end_id
  end

  public
  def convert
    @doc = Nokogiri::XML::Document.new
    @root = Nokogiri::XML::Node.new "testsuite",@doc
    @doc.add_child @root
    find_root_suites do |suite_id|
      @root.add_child(gen_suite(suite_id,1))
    end
    @doc.to_xml(:indent => 2)
  end

  def check
  end
end


if __FILE__ == $0
  f = File.open("data/all.xml.xls")
  ex=Excel2XMLHelper.new f
  str = ex.convert
  f = File.new("test.xml",'w')
  f.write(str)
  f.close
end
