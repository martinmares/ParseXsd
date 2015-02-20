# encoding: utf-8

require 'nokogiri'
require 'term/ansicolor'
require 'axlsx'
require 'trollop'

VERSION = "v0.1beta"
XMLSCHEMA = "http://www.w3.org/2001/XMLSchema"

# Parse ARGS~
opts = Trollop::options do
  version "parsexsd #{VERSION} (c) 2015 Martin Mareš"
  opt :xsd, 'name of the input XSD file', type: :string
  opt :xlsx, 'name of the output XLSX file', type: :string
  opt :stdout, 'write the XSD structure on the screen'
  opt :indent, 'name the elements in XLSX will be indented'
  opt :border, 'generate a border for cells in XLSX?'
  opt :columns, 'the list of columns in the XLSX', type: :string
  opt "request-end-with".to_sym, 'mark the line ending at {Request}', type: :string
  opt "response-end-with".to_sym, 'mark the line ending at {Response}', type: :string
  opt "header-request".to_sym, 'add a header to each of the {Request} elem.'
  opt "header-response".to_sym, 'add a header to each of the {Response} elem.'
  opt "auto-filter".to_sym, 'turn on the "auto filter on the first row"'
  opt "font-name".to_sym, 'change the default font ("Tahoma")', type: :string
  opt "font-size".to_sym, 'change the default font size (9)', type: :int
  opt "header-font-size".to_sym, 'change the default heder font size (9)', type: :int
end

config_file = opts[:config_file]
Trollop::die :xsd, "please specify input XSD file" unless opts[:xsd]
Trollop::die :xsd, "XSD file must exist" unless File.exist?(opts[:xsd]) if opts[:xsd]

@stdout = opts[:stdout] || false
@indent = opts[:indent] || false
@font = opts["font-name".to_sym] || 'Tahoma'
@font_size = opts["font-size".to_sym] || 9
@header_font_size = opts["header-font-size".to_sym] || 9
@xsd_file_name = opts[:xsd]
@xlsx_file_name = opts[:xlsx]
@auto_filter = opts["auto-filter".to_sym] || false
@border = opts[:border] || false

@test_request = opts["request-end-with".to_sym] || nil
@test_response = opts["response-end-with".to_sym] || nil

@header_request = opts["header-request".to_sym] || false
@header_response = opts["header-response".to_sym] || false

class String
  include Term::ANSIColor
end


Element = Struct.new(:name, :type, :ref, :its_complex_type, :its_simple_type, :min_occurs, :max_occurs, :nillable, :description, :deep, :inout, :its_recursion, :prefix)
Imported = Struct.new(:namespace, :schemalocation, :content)

@columns = {name: 'NAME', schematype: "XMLSCHEMA\nTYPE", type: 'TYPE', length: "LENGTH/\nPRECISION", multi: 'MULTIPL.', enum: "ENUM.\nVALUES", kind: 'KIND', desc: 'DESCRIPTION', mandatory: 'MANDATORY', complex: "COMPLEX\nTYPE", simple: "SIMPLE\nTYPE", minoccurs: "MIN\nOCCURS", maxoccurs: "MAX\nOCCURS", nill: 'NILLABLE'}

@header = Array.new
@empty_row = Array.new
@only_columns = Hash.new
@imported_schemas = Hash.new

# name schematype type length multi enum kind desc mandatory complex simple minoccurs maxoccurs nill
if opts[:columns]
  c = opts[:columns].split(',')
  if c
    c.each do |val|
      @only_columns[val.to_sym] = ''
    end
  end
else
  # if not exist --columns
  @columns.each do |key,val|
    @only_columns[key] = ''
  end
end

@columns.each do |key, val|
  @header << val if @only_columns[key]
  @empty_row << '' if @only_columns[key]
end

def check_recursion(node_types, actual_node, deep)
  #puts "#{padding(deep)}--> #{node_types} -- #{actual_node} -- "
  types = node_types.split(';') || []
  types.each do |value|
    return true if value == actual_node
  end
  false
end

xsd_file = opts[:xsd]

def load_xsd_file(filename)
  if File::exist? filename
    Nokogiri::XML(File.open(filename))
  else
    nil
  end
end

@doc = load_xsd_file(xsd_file)

# Load all namespaces
@namespaces = @doc.collect_namespaces
@xsd_namespace = ''
@namespaces.each do |key,val|
  @xsd_namespace = key if val == XMLSCHEMA
end
@xsd_prefix = @xsd_namespace.split(':')[1]

# Import XSD documents (xsd:import)
@import_elements = @doc.xpath('/namespace:schema/namespace:import', namespace: XMLSCHEMA)
unless @import_elements.nil?
  @import_elements.each do |elem|
    namespace = elem['namespace']
    schemalocation = elem['schemaLocation']
    # puts "Namespace: #{namespace}"
    # puts "Schema location: #{schemalocation}"
    @namespaces.each do |key,val|
      if val == namespace
        prefix = key.split(':')[1]
        # puts "Prefix: #{prefix}"
        content = load_xsd_file(schemalocation)
        @imported_schemas[prefix] = Imported.new(namespace, schemalocation, content) if content
        break
      end
    end
  end
end

# Search for root node
@root_node = @doc.xpath('/namespace:schema/namespace:element', namespace: XMLSCHEMA)
@elements = Hash.new
@enums = Hash.new
@cnt = 0

def padding(deep)
  " " * 4 * deep
end

def print_enums(node_name, enum_node, deep = 0)
  @enums[node_name] = ""
  enum_node.each do |node|
    puts "#{padding(deep)} * #{node['value']}" if @stdout
    @enums[node_name] += "#{node['value']}\n"
  end
  @enums[node_name].strip!
end

def exist_complex_type(doc, node, schema)
    complex_type_node = doc.xpath("/namespace:schema/namespace:complexType[@name='#{node['type']}']", namespace: schema)
    unless complex_type_node.nil?
      complex_type_node.size > 0
    else
      false
    end
end

def exist_simple_type(doc, node, schema)
    simple_type_node = doc.xpath("/namespace:schema/namespace:simpleType[@name='#{node['type']}']", namespace: schema)
    unless simple_type_node.nil?
      simple_type_node.size > 0
    else
      false
    end
end

def occurs(node)
  s = ""
  s = s + " (" + node['minOccurs'] unless node['minOccurs'].nil?
  s = s + " - " + node['maxOccurs'] unless node['maxOccurs'].nil?
  s = s + ")" unless s.empty?
  s.cyan
end

def nillable(node)
  s = ""
  unless node['nillable'].nil?
    nill = node['nillable']
    s = " " + "nillable(#{nill})".on_green if nill == 'true'
    s = " " + "nillable(#{nill})".on_red if nill == 'false'
  end
  s
end

def documentation(node, deep, schema)
  description_node = node.xpath("namespace:annotation/namespace:documentation", namespace: schema)
  if description_node.size > 0
    description = description_node[0].content
    puts "#{padding(deep)}# #{description}".magenta if @stdout && description
  end
  description || ''
end

def print_elements(doc, start_node, schema, deep = 0, inout = 'in', node_types = '', prefix = '')
  start_node.each do |node|
    key = "#{@cnt}-#{node['name']}-#{node['type']}-#{node['ref']}"
    io ||= inout
    io = 'out' if node['name'] && node['name'].end_with?(@test_response)
    if exist_complex_type(doc, node, schema)
      # Komplexní typ
      description = documentation(node, deep, schema)
      puts "#{padding(deep)}#{prefix}#{node['name']} #{node['type'].yellow}#{occurs(node)}#{nillable(node)} #{'@complexType'.on_blue} {deep: #{deep}}" if @stdout
      @elements[key] = Element.new(node['name'], node['type'], node['ref'], 'Y', '', node['minOccurs'], node['maxOccurs'], node['nillable'], description, deep, io, false, prefix)
      complex_type_node = doc.xpath("/namespace:schema/namespace:complexType[@name='#{node['type']}']", namespace: schema)
      extension_base = complex_type_node.xpath("descendant::*/namespace:extension", namespace: schema)
      # Extension
      if extension_base.size > 0
        base_name = extension_base[0]['base']
        puts "#{padding(deep + 1)}#{prefix}#{base_name.on_yellow} #{'@extension'.on_blue}" if @stdout
        extension_base_node = doc.xpath("/namespace:schema/namespace:complexType[@name='#{base_name}']", namespace: schema)
        extension_element_nodes = extension_base_node.xpath("descendant::*/namespace:element | descendant::*/namespace:group", namespace: schema)
        # Test for prefix, e.g. "prefix:complexType"...
        # if success, print imported elements
        if base_name.split(':').size == 2
          imp_prefix = base_name.split(':')[0]
          imp_type = base_name.split(':')[1]
          imp_doc = @imported_schemas[imp_prefix][:content]
          imp_complex_type_node = imp_doc.xpath("/namespace:schema/namespace:complexType[@name='#{imp_type}']", namespace: schema)
          imp_element_nodes = imp_complex_type_node.xpath("descendant::*/namespace:element | descendant::*/namespace:group", namespace: schema)
          # p imp_type
          # p imp_complex_type_node.size
          print_elements(imp_doc, imp_element_nodes, schema, deep + 2, io, '', "#{imp_prefix}:") unless imp_complex_type_node.empty?
        end
        print_elements(doc, extension_element_nodes, schema, deep + 2, io, '', prefix) unless extension_base_node.empty?
      end
      element_nodes = complex_type_node.xpath("descendant::*/namespace:element | descendant::*/namespace:group", namespace: schema)
      unless check_recursion(node_types, "#{node['type']}", deep)
        print_elements(doc, element_nodes, schema, deep + 1, io, "#{node_types}#{node['type']};", prefix) unless complex_type_node.empty?
      else
        @elements[key][:description] = "Rekurze komplexního typu \"#{node['type']}\"..."
        @elements[key][:its_recursion] = true
        return
      end
    elsif exist_simple_type(doc, node, schema)
      # Výčet (SimpleType)
      description = documentation(node, deep, schema)
      puts "#{padding(deep)}#{prefix}#{node['name']} #{node['type'].on_magenta}#{occurs(node)}#{nillable(node)} #{'@simpleType'.on_blue}" if @stdout
      @elements[key] = Element.new(node['name'], node['type'], node['ref'], '', 'Y', node['minOccurs'], node['maxOccurs'], node['nillable'], description, deep, io, false, prefix)
      simple_type_node = doc.xpath("/namespace:schema/namespace:simpleType[@name='#{node['type']}']", namespace: schema)
      enum_nodes = simple_type_node.xpath("descendant::*/namespace:enumeration", namespace: schema)
      print_enums(node['type'], enum_nodes, deep + 1) unless simple_type_node.empty?
    else
      # Element
      description = documentation(node, deep, schema)
      puts "#{padding(deep)}#{prefix}#{node['name']} [#{node['type'].green}]#{occurs(node)}#{nillable(node)} #{'@element'.on_blue}" unless node['name'].nil? if @stdout
      @elements[key] = Element.new(node['name'], node['type'], node['ref'], '', '', node['minOccurs'], node['maxOccurs'], node['nillable'], description, deep, io, false, prefix)
      unless node['ref'].nil?
        ref_type_node = doc.xpath("/namespace:schema/namespace:group[@name='#{node['ref']}']", namespace: schema)
        group_nodes = ref_type_node.xpath("descendant::*/namespace:element | descendant::*/namespace:group", namespace: schema)
        print_elements(doc, group_nodes, schema, deep + 1, io, '', prefix) unless ref_type_node.empty?
      end
    end
    @cnt += 1
  end
end

def arrows(deep)
  if @indent
    '  ' * deep
  else
    ''
  end
end

def get_row_data(row_data)
  row = Array.new
  row_data.each do |key, val|
    row << val if @only_columns[key]
  end
  row
end

# XLS
def save_xlsx
  puts "\n#{'=' * 10} Generating XLSX #{'=' * 13}"
  p = Axlsx::Package.new
  p.use_autowidth = true
  wb = p.workbook
  wb.styles do |s|
    # Excel 2007/2010 Indexed Colors
    # https://closedxml.codeplex.com/wikipage?title=Excel%20Indexed%20Colors
    border = (@border ? Axlsx::STYLE_THIN_BORDER : 0)
    head = s.add_style font_name: @font, sz: @header_font_size, family: 1, b: false, fg_color: "FF000000", bg_color: "FFC0C0C0",
      :alignment => { :horizontal => :left, :vertical => :top, :wrap_text => true }, border: border
    null_cell = s.add_style font_name: @font, sz: @font_size, family: 1,
      :alignment => { :horizontal => :left, :vertical => :top, :wrap_text => true }
    normal_cell = s.add_style font_name: @font, sz: @font_size, family: 1,
      :alignment => { :horizontal => :left, :vertical => :top, :wrap_text => true }, border: border
    italic_cell = s.add_style font_name: @font, sz: @font_size, family: 1, i: true, b: true,
      :alignment => { :horizontal => :left, :vertical => :top, :wrap_text => true }, border: border
    complex_cell = s.add_style font_name: @font, sz: @font_size, family: 1, b: true,
      :alignment => { :horizontal => :left, :vertical => :top, :wrap_text => true }, border: border
    enum_cell = s.add_style fg_color: "FF0000FF", font_name: @font, sz: @font_size, family: 1, b: true,
      :alignment => { :horizontal => :left, :vertical => :top, :wrap_text => true }, border: border
    marked_cell = s.add_style font_name: @font, sz: @font_size, family: 1, b: true, fg_color: "FF000000", bg_color: "FFF0F0F0",
      :alignment => { :horizontal => :left, :vertical => :top, :wrap_text => true }, border: border
    recursion_cell = s.add_style font_name: @font, sz: @font_size, family: 1, b: true, fg_color: "FF0000FF", bg_color: "FFFFFFCC",
      :alignment => { :horizontal => :left, :vertical => :top, :wrap_text => true }, border: border
    wb.add_worksheet(:name => "XSD") do |sheet|

      sheet.add_row @header

      i = 0
      @elements.each do |key, val|
        struct = @elements[key]
        name = struct[:name]
        type = struct[:type]
        values = ''
        mandatory = ''
        multiplicity = ''

        if struct[:its_complex_type] == 'Y'
          name = "#{struct[:name]}"
          basic_type = "STRUCT\n(#{struct[:type]})"
          type = "ComplexType\n(#{struct[:type]})"
        elsif struct[:its_simple_type] == 'Y'
          values = @enums[struct[:type]]
          basic_type = "ENUM\n(#{struct[:type]})"
          type = "SimpleType\n(#{struct[:type]})"
        elsif not struct[:ref].to_s.empty?
          name = "#{struct[:ref]}"
          basic_type = "GROUP"
          type = "Reference\n(#{struct[:ref]})"
        end

        case type
        when "#{@xsd_prefix}:int"
          basic_type = 'INTEGER'
        when "#{@xsd_prefix}:integer"
          basic_type = 'INTEGER'
        when "#{@xsd_prefix}:double"
          basic_type = 'DOUBLE'
        when "#{@xsd_prefix}:base64Binary"
          basic_type = 'BASE64'
        when "#{@xsd_prefix}:date"
          basic_type = 'DATE'
        when "#{@xsd_prefix}:dateTime"
          basic_type = 'DATE+TIME'
        when "#{@xsd_prefix}:string"
          basic_type = 'STRING'
        end

        min_occurs = struct[:min_occurs] || '1'
        max_occurs = struct[:max_occurs] || '1'

        mandatory = 'Y' if min_occurs == '1' && max_occurs == '1'
        mandatory = 'O' if min_occurs == '0'

        multiplicity = '0..1' if min_occurs == '0' && max_occurs == '1'
        multiplicity = '0..N' if min_occurs == '0' && max_occurs == 'unbounded'
        multiplicity = '1..N' if min_occurs == '1' && max_occurs == 'unbounded'
        multiplicity = '1' if min_occurs == '1' && max_occurs == '1'

        if @header_request && name.end_with?(@test_request) && i != 0
          sheet.add_row @empty_row
          i += 1
          sheet.row_style i, null_cell
          sheet.add_row @header
          i += 1
          sheet.row_style i, head
        end

        if @header_response && name.end_with?(@test_response) && i != 0
          sheet.add_row @empty_row
          i += 1
          sheet.row_style i, null_cell
          sheet.add_row @header
          i += 1
          sheet.row_style i, head
        end

        desc = struct[:description]
        length = desc.scan(/\$length\(\d+\)/)
        format = desc.scan(/\$format\(\S+\)/)
        desc = desc.gsub(/\$length(.*)|\$format(.*)/, '').strip

        length_prec_format = ''
        length_prec_format = length[0].scan(/\d+/)[0] if (length && length.size > 0)
        length_prec_format = format[0].scan(/\(\S+\)/)[0][1..-2] if (format && format.size > 0)

        # name schematype type length multi enum kind desc mandatory complex simple minoccurs maxoccurs nill
        row_data = {name: "#{arrows(struct[:deep])}#{name}",
               schematype: type,
               type: basic_type,
               length: length_prec_format,
               multi: multiplicity,
               enum: values,
               kind: struct[:inout],
               desc: desc,
               mandatory: mandatory,
               complex: struct[:its_complex_type],
               simple: struct[:its_simple_type],
               minoccurs: min_occurs,
               maxoccurs: max_occurs,
               nill: struct[:nillable]
        }

        row = get_row_data(row_data)

        style = normal_cell
        style = italic_cell unless struct[:ref].nil?
        style = complex_cell if struct[:its_complex_type] == 'Y'
        style = enum_cell if struct[:its_simple_type] == 'Y'
        style = recursion_cell if struct[:its_recursion]
        style = marked_cell if name.end_with? @test_request
        style = marked_cell if name.end_with? @test_response

        #unless struct[:prefix].empty?
        #  style << s.add_style 
        #end

        sheet.add_row row, style: style
        #sheet.row_style i+1, style

        i += 1
      end
      sheet.auto_filter = "A1:H1" if @auto_filter
      sheet.row_style 0, head
    end
  end
  puts "Save to file #{@xlsx_file_name}"
  p.serialize(@xlsx_file_name)
  puts "Count of elements: #{@elements.size}"
  puts "#{'=' * 40}"
end

print_elements(@doc, @root_node, XMLSCHEMA) if @xsd_file_name

# Print
puts "\n#{'=' * 10} List of namespaces #{'=' * 10}"
puts "Default XMLSchema namespace is '#{@xsd_namespace}'"
@namespaces.each do |key,val|
  puts "#{key} - #{val}"
end
puts "#{'=' * 40}"

save_xlsx() if @xlsx_file_name

