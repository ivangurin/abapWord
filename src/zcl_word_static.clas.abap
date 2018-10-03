*----------------------------------------------------------------------*
"       CLASS ZCL_WORD_STATIC DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
class zcl_word_static definition
  public
  final
  create public .

  public section.

    types:
      begin of ty_border,
        " Border width
        " Integer: 8 = 1pt
        " Default: no border
        width type i,
        " Space between border and content
        " Integer
        " Default: document default
        space type i,
        " Border color
        " You can use the predefined font color constants or specify any rgb hexa color code
        " Default : document default
        color type string,
        " Border style
        " You can use the predefined border style constants
        " Default : document default
        style type string, "border style
      end of ty_border .
    types:
      begin of ts_numbering,
        id      type string,
        level   type string,
        restart type abap_bool,
      end of ts_numbering .
    types:
      begin of ty_character_style_effect,
        " Font name to use for the character text fragment
        " You can use the predefined font constants
        " Default : document default
        font      type string,
        " Size of the font in pt
        " Default : document default
        size      type i,
        " Font color to apply to the character text fragment
        " You can use the predefined font color constants or specify any rgb hexa color code
        " Default : document default
        color     type string,
        " Background color to apply to the character text fragment
        " You can use the predefined font color constants or specify any rgb hexa color code
        " Default : document default
        bgcolor   type string,
        " Highlight color to apply to the character text fragment
        " You must use the predefined highlight color constants (limited color choice)
        " If you want to use other color, please use bgcolor instead of highlight
        " Default : document default
        highlight type string,
        " Set character text fragment as bold (boolean)
        " You must use predefined true/false constants
        " Default : not bold
        bold      type i,
        " Set character text fragment as italic (boolean)
        " You must use predefined true/false constants
        " Default : not italic
        italic    type i,
        " Set character text fragment as underline (boolean)
        " You must use predefined true/false constants
        " Default : not underline
        underline type i,
        " Set character text fragment as strike (boolean)
        " You must use predefined true/false constants
        " Default : not strike
        strike    type i,
        " Set character text fragment as exponent (boolean)
        " You must use predefined true/false constants
        " Default : not exponent
        sup       type i,
        " Set character text fragment as subscript (boolean)
        " You must use predefined true/false constants
        " Default : not subscript
        sub       type i,
        " Set character text fragment as upper case (boolean)
        " You must use predefined true/false constants
        " Default : not upper case
        caps      type i,
        " Set character text fragment as small upper case (boolean)
        " You must use predefined true/false constants
        " Default : not small upper case
        smallcaps type i,
        " Letter spacing to apply to the character text fragment
        " 0 = normal, +20 = expand 1pt, -20 = condense 1pt
        " Default : document default
        spacing   type string,
        " Name of the label to use for this text fragment
        " Be carefull that if c_label_anchor is not found in text fragment,
        " this attribute is ignored
        label     type string,
      end of ty_character_style_effect .
    types:
      begin of ty_paragraph_style_effect,
        " Set alignment for the paragraph (left, right, center, justify)
        " Use the predefined alignment constants
        " Default : document default
        alignment           type string,
        " Set spacing before paragraph to "auto" (boolean)
        " Use the predefined true/false constants
        " Default : document default
        spacing_before_auto type i, "boolean
        " Set spacing before paragraph
        " Integer value, 20 = 1pt
        " Default : document default
        spacing_before      type string,
        " Set spacing after paragraph to "auto" (boolean)
        " Use the predefined true/false constants
        " Default : document default
        spacing_after_auto  type i, "boolean
        " Set spacing after paragraph
        " Integer value: 20 = 1pt
        " Default : document default
        spacing_after       type string,
        " Set interline in paragraph
        " Integer value: 240 = normal interline, 120 = multiple x0.5, 480 = multiple x2
        " Default : document default
        interline           type i,
        " Set left indentation in paragraph
        " Integer value: 567 = 1cm
        " Default : document default
        leftindent          type string,
        " Set right indentation in paragraph
        " Integer value: 567 = 1cm
        " Default : document default
        rightindent         type string,
        " Set left indentation for first line in paragraph
        " Integer value: 567 = 1cm. Negative value allowed
        " Default : document default
        firstindent         type string,
        " Add a breakpage before paragraph (boolean)
        " Use the predefined true/false constants
        " Default : no breakpage
        break_before        type i,
        " Set the hierarchical title level of the paragraph
        " 1 for title1, 2 for title2...
        " Default : not a hierarchical title
        hierarchy_level     type i,
        " Set the left border of the paragraph
        " See class type ty_border for details
        " Default : No border
        border_left         type ty_border,
        " Set the top border of the paragraph
        " See class type ty_border for details
        " Default : No border
        border_top          type ty_border,
        " Set the right border of the paragraph
        " See class type ty_border for details
        " Default : No border
        border_right        type ty_border,
        " Set the bottom border of the paragraph
        " See class type ty_border for details
        " Default : No border
        border_bottom       type ty_border,
        " Background color to apply to the paragraph
        " You can use the predefined font color constants or specify any rgb hexa color code
        " Default : document default
        bgcolor             type string,
        " Numbering
        numbering           type ts_numbering,
      end of ty_paragraph_style_effect .
    types:
      begin of ty_list_style,
        type type string,
        name type string,
      end of ty_list_style .
    types:
      ty_list_style_table type standard table of ty_list_style .
    types:
      begin of ty_list_object,
        id   type string,
        type type string,
        path type string,
      end of ty_list_object .
    types:
      ty_list_object_table type standard table of ty_list_object .
    types:
      begin of ty_table_style_field,
        " Content of the cell
        text              type string,
        " Character style name
        style             type string,
        " Direct character style effect
        style_effect      type ty_character_style_effect,
        " Paragraph style name
        line_style        type string,
        " Direct paragraph style effect
        line_style_effect type ty_paragraph_style_effect,
        " Cell background color in hexa RGB
        bgcolor           type string,
        " Set vertical alignment for cell
        valign            type string,
        " Set number of horizontal cell merged
        " Start from 2 to merge the next cell with current
        " Next cell will be completely ignored
        " 0 or 1 to ignore this parameter
        merge             type i,
        " start
        " continue
        vmerge            type string,
        " Instead of text, insert an image in the cell
        " textline, style, style_effect are ignored
        image_id          type string,
        image_width       type i,
        image_height      type i,
        " For complex cell content, insert xml paragraph fragment
        " You cannot insert xml fragment that are not in a (or many) paragraph
        " textline, style, style_effect, line_style, line_style_effect, image_id are ignored
        xml               type string,
        width             type string,
        direction         type string,
        hide              type abap_bool,
      end of ty_table_style_field .
    types:
      begin of ty_style_table,
        firstcol type i, "boolean, first col is different
        firstrow type i, "boolean, first row is different
        lastcol  type i, "boolean, last col is different
        lastrow  type i, "boolean, last row is different
        nozebra  type i, "boolean, no line zebra
        novband  type i, "boolean, no column zebra
      end of ty_style_table .

    constants c_true type i value 1.                        "#EC NOTEXT
    constants c_false type i value 0.                       "#EC NOTEXT
    " Predefined fields that can be added
    constants c_field_pagecount type string value '##FIELD#PAGE##'. "#EC NOTEXT                                                                                          " MS Word current page counter
    constants c_field_pagetotal type string value '##FIELD#NUMPAGES##'. "#EC NOTEXT                                                                                          " MS Word total page counter
    constants c_field_title type string value '##FIELD#TITLE##'. "#EC NOTEXT
    constants c_field_author type string value '##FIELD#AUTHOR##'. "#EC NOTEXT
    constants c_field_authorlastmod type string value '##FIELD#LASTSAVEDBY##'. "#EC NOTEXT
    constants c_field_comments type string value '##FIELD#COMMENTS##'. "#EC NOTEXT
    constants c_field_keywords type string value '##FIELD#KEYWORDS##'. "#EC NOTEXT
    constants c_field_category type string value '##FIELD#DOCPROPERTY CATEGORY##'. "#EC NOTEXT
    constants c_field_subject type string value '##FIELD#SUBJECT##'. "#EC NOTEXT
    constants c_field_revision type string value '##FIELD#REVNUM##'. "#EC NOTEXT
    constants c_field_creationdate type string value '##FIELD#CREATEDATE##'. "#EC NOTEXT
    constants c_field_moddate type string value '##FIELD#SAVEDATE##'. "#EC NOTEXT
    constants c_field_todaydate type string value '##FIELD#DATE##'. "#EC NOTEXT
    constants c_field_filename type string value '##FIELD#FILENAME##'. "#EC NOTEXT
    " Anchor for label
    constants c_label_anchor type string value '##LABEL##'. "#EC NOTEXT
    constants:
      " Prefix for SAPWR url
      c_sapwr_prefix(6) type c value 'SAPWR:'.              "#EC NOTEXT
    " Predefined fonts
    constants c_font_arial type string value 'Arial'.       "#EC NOTEXT
    constants c_font_times type string value 'Times New Roman'. "#EC NOTEXT
    constants c_font_comic type string value 'Comic Sans MS'. "#EC NOTEXT
    constants c_font_calibri type string value 'Calibri'.   "#EC NOTEXT
    constants c_font_cambria type string value 'Cambria'.   "#EC NOTEXT
    constants c_font_courier type string value 'Courier New'. "#EC NOTEXT
    constants c_font_symbol type string value 'Wingdings'.  "#EC NOTEXT
    " Predefined Styles
    " Regarding your document language, style could be named differently
    " For example, in french ms document, title1 is labelled "Titre1"
    constants c_style_title type string value 'Title'.      "#EC NOTEXT
    constants c_style_heading1 type string value 'Heading1'. "#EC NOTEXT
    constants c_style_heading2 type string value 'Heading2'. "#EC NOTEXT
    constants c_style_heading3 type string value 'Heading3'. "#EC NOTEXT
    constants c_style_heading4 type string value 'Heading4'. "#EC NOTEXT
    constants c_style_normal type string value 'Normal'.    "#EC NOTEXT
    constants c_style_no_spacing type string value 'NoSpacing'. "#EC NOTEXT
    " Predefined break types
    constants c_breaktype_line type i value 6.              "#EC NOTEXT
    constants c_breaktype_page type i value 7.              "#EC NOTEXT
    constants c_breaktype_column type i value 1. "#EC NOTEXT                                                                                          "NOT MANAGED
    constants c_breaktype_section type i value 2.           "#EC NOTEXT
    constants c_breaktype_section_continuous type i value 3. "#EC NOTEXT
    " Predefined symbols
    constants c_symbol_checkbox_checked type string value 'ю'. "#EC NOTEXT
    constants c_symbol_checkbox type string value 'o'.      "#EC NOTEXT
    " Draw objects
    constants c_draw_image type i value 0.                  "#EC NOTEXT
    constants c_draw_rectangle type i value 1.              "#EC NOTEXT
    " Text alignment
    constants c_align_left type string value 'left'.        "#EC NOTEXT
    constants c_align_center type string value 'center'.    "#EC NOTEXT
    constants c_align_right type string value 'right'.      "#EC NOTEXT
    constants c_align_justify type string value 'both'.     "#EC NOTEXT
    " Vertical alignment
    constants c_valign_top type string value 'top'.         "#EC NOTEXT
    constants c_valign_middle type string value 'center'.   "#EC NOTEXT
    constants c_valign_bottom type string value 'bottom'.   "#EC NOTEXT
    " Types for header/footer
    constants c_type_header type string value 'header'.     "#EC NOTEXT
    constants c_type_footer type string value 'footer'.     "#EC NOTEXT
    " Style type
    constants c_type_paragraph type string value 'paragraph'. "#EC NOTEXT
    constants c_type_character type string value 'character'. "#EC NOTEXT
    constants c_type_table type string value 'table'.       "#EC NOTEXT
    constants c_type_numbering type string value 'numbering'. "#EC NOTEXT
    " Image type
    constants c_type_image type string value 'image'.       "#EC NOTEXT
    " Page orientation
    constants c_orient_landscape type i value 1. "#EC NOTEXT                                                                                          " Landscape
    constants c_orient_portrait type i value 0. "#EC NOTEXT                                                                                          " Portrait
    " Note type
    constants c_notetype_foot type i value 1. "#EC NOTEXT                                                                                          "foot note
    constants c_notetype_end type i value 2. "#EC NOTEXT                                                                                          "end note
    " Predefined colors
    constants c_color_black type string value '000000'.     "#EC NOTEXT
    constants c_color_blue type string value '0000FF'.      "#EC NOTEXT
    constants c_color_turquoise type string value '00FFFF'. "#EC NOTEXT
    constants c_color_brightgreen type string value '00FF00'. "#EC NOTEXT
    constants c_color_pink type string value 'FF00FF'.      "#EC NOTEXT
    constants c_color_red type string value 'FF0000'.       "#EC NOTEXT
    constants c_color_yellow type string value 'FFFF00'.    "#EC NOTEXT
    constants c_color_white type string value 'FFFFFF'.     "#EC NOTEXT
    constants c_color_darkblue type string value '000080'.  "#EC NOTEXT
    constants c_color_teal type string value '008080'.      "#EC NOTEXT
    constants c_color_green type string value '008000'.     "#EC NOTEXT
    constants c_color_violet type string value '800080'.    "#EC NOTEXT
    constants c_color_darkred type string value '800000'.   "#EC NOTEXT
    constants c_color_darkyellow type string value '808000'. "#EC NOTEXT
    constants c_color_gray type string value '808080'.      "#EC NOTEXT
    constants c_color_lightgray type string value 'C0C0C0'. "#EC NOTEXT
    " Predefined border style
    constants c_border_simple type string value 'single'.   "#EC NOTEXT
    constants c_border_double type string value 'double'.   "#EC NOTEXT
    constants c_border_triple type string value 'triple'.   "#EC NOTEXT
    constants c_border_dot type string value 'dotted'.      "#EC NOTEXT
    constants c_border_dash type string value 'dashed'.     "#EC NOTEXT
    constants c_border_wave type string value 'wave'.       "#EC NOTEXT
    " Predefined font highlight color
    constants c_highlight_yellow type string value 'yellow'. "#EC NOTEXT
    constants c_highlight_green type string value 'green'.  "#EC NOTEXT
    constants c_highlight_cyan type string value 'cyan'.    "#EC NOTEXT
    constants c_highlight_magenta type string value 'magenta'. "#EC NOTEXT
    constants c_highlight_blue type string value 'blue'.    "#EC NOTEXT
    constants c_highlight_red type string value 'red'.      "#EC NOTEXT
    constants c_highlight_darkblue type string value 'darkBlue'. "#EC NOTEXT
    constants c_highlight_darkcyan type string value 'darkCyan'. "#EC NOTEXT
    constants c_highlight_darkgreen type string value 'darkGreen'. "#EC NOTEXT
    constants c_highlight_darkmagenta type string value 'darkMagenta'. "#EC NOTEXT
    constants c_highlight_darkred type string value 'darkRed'. "#EC NOTEXT
    constants c_highlight_darkyellow type string value 'darkYellow'. "#EC NOTEXT
    constants c_highlight_darkgray type string value 'darkGray'. "#EC NOTEXT
    constants c_highlight_lightgray type string value 'lightGray'. "#EC NOTEXT
    constants c_highlight_black type string value 'black'.  "#EC NOTEXT
    constants c_basesize type i value 12700.                "#EC NOTEXT
    class-data pr_doc_id type i value 100.                  "#EC NOTEXT

    class zcl_word_static definition load .
    class-methods build_text_style
      importing
        !i_style     type simple optional
        !is_style    type zcl_word_static=>ty_character_style_effect optional
      returning
        value(e_xml) type string .
    class-methods build_text
      importing
        !i_text      type simple
        !i_style     type simple optional
        !is_style    type zcl_word_static=>ty_character_style_effect optional
      returning
        value(e_xml) type string .
    class-methods build_paragraph_style
      importing
        !i_style     type simple optional
        !is_style    type zcl_word_static=>ty_paragraph_style_effect optional
        !i_section   type simple optional
      returning
        value(e_xml) type string .
    class-methods build_paragraph
      importing
        !i_text      type simple optional
        !i_style     type simple optional
        !is_style    type zcl_word_static=>ty_paragraph_style_effect optional
        !i_section   type simple optional
          preferred parameter i_text
      returning
        value(e_xml) type string .
    class-methods build_abstract_num
      importing
        !i_id        type simple
      returning
        value(e_xml) type string .
    class-methods build_num
      importing
        !i_abstract_num_id type simple
        !i_id              type simple
      returning
        value(e_xml)       type string .
    class-methods build_container
      importing
        !i_id        type simple
        !i_xml       type simple optional
      returning
        value(e_xml) type string .
    type-pools abap .
    class-methods build_image
      importing
        !i_id        type simple
        !i_width     type i
        !i_height    type i
        !i_zoom      type f optional
        !i_landscape type abap_bool default abap_false
      returning
        value(e_xml) type string .
    class-methods build_table
      importing
        !it_data     type standard table
        !i_style     type simple optional
        !is_style    type zcl_word_static=>ty_style_table optional
        !i_border    type i default zcl_word_static=>c_true
        !i_width     type simple optional
        !i_indent    type i optional
      returning
        value(e_xml) type string .
    class-methods protect_string
      importing
        !in        type simple
      returning
        value(out) type string .
    class-methods protect_label
      importing
        !in        type simple
      returning
        value(out) type string .
    class-methods get_containers_id
      importing
        !i_xml       type simple
      returning
        value(et_id) type stringtab .
    class-methods set_container_content
      importing
        !i_id  type simple
        !i_xml type simple optional
      changing
        !c_xml type simple .
    class-methods get_container_content
      importing
        !i_id        type simple
        !i_xml       type simple
      returning
        value(e_xml) type string .
    class-methods get_variables
      importing
        !i_xml         type simple
      returning
        value(et_list) type stringtab .
  protected section.
  private section.
ENDCLASS.



CLASS ZCL_WORD_STATIC IMPLEMENTATION.


  method build_abstract_num.

    data l_guid type guid_32.
    l_guid = zcl_abap_static=>create_guid( ).

    e_xml =
      '<w:abstractNum w:abstractNumId="' && zcl_abap_static=>value2text( i_id ) && '">' &&
        '<w:nsid w:val="' && l_guid+24(8) && '"/>' &&
        '<w:multiLevelType w:val="hybridMultilevel"/>' &&
        '<w:lvl w:ilvl="0">' &&
          '<w:start w:val="1"/>' &&
          '<w:numFmt w:val="decimal"/>' &&
          '<w:lvlText w:val="%1."/>' &&
          '<w:lvlJc w:val="left"/>' &&
          '<w:pPr>' &&
            '<w:ind w:left="720" w:hanging="360"/>' &&
          '</w:pPr>' &&
        '</w:lvl>' &&
        '<w:lvl w:ilvl="1">' &&
          '<w:start w:val="1"/>' &&
          '<w:numFmt w:val="lowerLetter"/>' &&
          '<w:lvlText w:val="%2."/>' &&
          '<w:lvlJc w:val="left"/>' &&
          '<w:pPr>' &&
            '<w:ind w:left="1420" w:hanging="360"/>' &&
          '</w:pPr>' &&
        '</w:lvl>' &&
        '<w:lvl w:ilvl="2">' &&
          '<w:start w:val="1"/>' &&
          '<w:numFmt w:val="lowerRoman"/>' &&
          '<w:lvlText w:val="%3."/>' &&
          '<w:lvlJc w:val="left"/>' &&
          '<w:pPr>' &&
            '<w:ind w:left="2160" w:hanging="360"/>' &&
          '</w:pPr>' &&
        '</w:lvl>' &&
      '</w:abstractNum>'.

  endmethod.


  method build_container.

    e_xml =
      '<w:sdt>' &&
        '<w:sdtPr>' &&
          '<w:tag w:val="' && i_id && '"/>' &&
          '<w15:appearance w15:val="hidden"/>' &&
        '</w:sdtPr>' &&
        '<w:sdtContent>' &&
        i_xml &&
        '</w:sdtContent>' &&
      '</w:sdt>'.

  endmethod.


  method build_image.

    " Calculate image scale
    if i_landscape eq abap_true.

      " Max X in landscape : 8877300
      " Max Y in landscape : 5743575 "could be less... depend of header/footer
***      l_scalex = 8877300 / l_width.
***      l_scaley = 5762625 / l_hight.

      data l_scalex type f.
      l_scalex = 11693190 / i_width.

      data l_scaley type f.
      l_scaley = 7582401 / i_height.

    else.

      " Max X in portrait : 5762625
      " Max Y in portrait : 8886825 "could be less... depend of header/footer
***      l_scalex = 5762625 / l_width.
***      l_scaley = 8886825 / l_hight.

      l_scalex = 7582401 / i_width.

      l_scaley = 11693190 / i_height.

    endif.

    if l_scalex < l_scaley.
      data l_scale_max type i.
      l_scale_max = floor( l_scalex ).
    else.
      l_scale_max = floor( l_scaley ).
    endif.

    " Image is smaler than page
    if l_scale_max gt 9525.

      data l_scale type i.
      l_scale = 9525. "no zoom

      if i_zoom ne 0.

        l_scale = l_scale * i_zoom.

        if l_scale gt l_scale_max.
          l_scale = l_scale_max.
        endif.

      endif.

    else.

      " Image is greater than page
      l_scale = l_scale_max.

      " if zoom is supplied and zoom < 1.
      if i_zoom ne 0.

        l_scale_max = 9525 * i_zoom.

        if l_scale_max lt l_scale.
          l_scale = l_scale_max.
        endif.

      endif.

    endif.

    data l_x type i.
    l_x = i_width * l_scale.

    data l_y type i.
    l_y = i_height * l_scale.

***    l_x = l_width * 9525.
***    l_y = l_hight * 9525.

    add 1 to pr_doc_id.

    data l_pr_doc_id type string.
    l_pr_doc_id = zcl_abap_static=>value2text( pr_doc_id ).

    data l_cx type string.
    l_cx = zcl_abap_static=>value2text( l_x ).

    data l_cy type string.
    l_cy = zcl_abap_static=>value2text( l_y ).

    concatenate
      '<w:r>'
      '<w:drawing>'

      '<wp:anchor allowOverlap="1" layoutInCell="1" locked="0" behindDoc="1" relativeHeight="251658240" simplePos="0" distR="0" distL="0" distB="0" distT="0">'
      '<wp:simplePos y="0" x="0"/>'
      '<wp:positionH relativeFrom="column">'
      '<wp:posOffset>0</wp:posOffset>'
      '</wp:positionH>'
      '<wp:positionV relativeFrom="paragraph">'
      '<wp:posOffset>0</wp:posOffset>'
      '</wp:positionV>'
      '<wp:extent cx="' l_cx '" cy="' l_cy '"/>'
      '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
      '<wp:wrapTopAndBottom/>'

*  **    '<wp:inline distT="0" distB="0" distL="0" distR="0">'
*  **    '<wp:extent cx="' l_x_string '" cy="' l_y_string '"/>'

      '<wp:docPr id="' l_pr_doc_id '" name=""/>'
      '<wp:cNvGraphicFramePr/>'
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
      '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
      '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
      '<pic:nvPicPr>'
      '<pic:cNvPr id="0" name=""/>'
      '<pic:cNvPicPr/>'
      '</pic:nvPicPr>'
      '<pic:blipFill>'
      '<a:blip r:embed="'
      i_id
      '"/>'
      '<a:stretch>'
      '<a:fillRect/>'
      '</a:stretch>'
      '</pic:blipFill>'
      '<pic:spPr>'
      '<a:xfrm>'
      '<a:off x="0" y="0"/>'
      '<a:ext cx="'
      l_cx
      '" cy="'
      l_cy
      '"/>'
      '</a:xfrm>'
      '<a:prstGeom prst="rect">'
      '<a:avLst/>'
      '</a:prstGeom>'
      '</pic:spPr>'
      '</pic:pic>'
      '</a:graphicData>'
      '</a:graphic>'

*  **    '</wp:inline>'

      '</wp:anchor>'

      '</w:drawing>'
      '</w:r>'
      into e_xml.

  endmethod.


  method build_num.

    e_xml =
      '<w:num w:numId="' && zcl_abap_static=>value2text( i_id ) && '">'  &&
        '<w:abstractNumId w:val="' && zcl_abap_static=>value2text( i_abstract_num_id ) && '"/>'  &&
      '</w:num>'.

  endmethod.


  method build_paragraph.

    data l_style type string.
    l_style =
      zcl_word_static=>build_paragraph_style(
        i_style   = i_style
        is_style  = is_style
        i_section = i_section ).

    e_xml = '<w:p>' &&  l_style && i_text && '</w:p>'.

  endmethod.                    "write_line


  method build_paragraph_style.

    if i_style is not initial.
      e_xml = '<w:pStyle w:val="' && i_style && '"/>'.
    endif.

    if is_style-break_before = zcl_word_static=>c_true.
      concatenate e_xml
                  '<w:pageBreakBefore/>'
                  into e_xml.
    endif.

    if not is_style-hierarchy_level is initial.
      data l_size type string.
      l_size = is_style-hierarchy_level - 1.
      condense l_size no-gaps.
      concatenate e_xml
                  '<w:outlineLvl w:val="'
                  l_size
                  '"/>'
                  into e_xml.
    endif.

    if not is_style-alignment is initial.
      concatenate e_xml
                  '<w:jc w:val="'
                  is_style-alignment
                  '"/>'
                  into e_xml.
    endif.

    if not is_style-bgcolor is initial.
      concatenate e_xml
                  '<w:shd w:val="clear" w:color="auto" w:fill="'
                  is_style-bgcolor
                  '"/>'
                  into e_xml.
    endif.

    data l_substyle type string.
    clear l_substyle.
    if is_style-spacing_before_auto = zcl_word_static=>c_true.
      l_substyle = ' w:beforeAutospacing="1"'.
    elseif not is_style-spacing_before is initial and is_style-spacing_before co '0123456789 '.
      data l_space type string.
      l_space = is_style-spacing_before.
      condense l_space no-gaps.
      concatenate l_substyle
                  ' w:beforeAutospacing="0" w:before="'
                  l_space
                  '"'
                  into l_substyle respecting blanks.
    endif.

    if is_style-spacing_after_auto = zcl_word_static=>c_true.
      concatenate l_substyle
                  ' w:afterAutospacing="1"'
                  into l_substyle respecting blanks.
    elseif not is_style-spacing_after is initial and is_style-spacing_after co '0123456789 '.
      l_space = is_style-spacing_after.
      condense l_space no-gaps.
      concatenate l_substyle
                  ' w:afterAutospacing="0" w:after="'
                  l_space
                  '"'
                  into l_substyle respecting blanks.
    endif.

    if not is_style-interline is initial.
      l_space = is_style-interline.
      condense l_space no-gaps.
      concatenate l_substyle
                  ' w:line="'
                  l_space
                  '"'
                  into l_substyle respecting blanks.
    endif.

    if not l_substyle is initial.
      concatenate e_xml
                  '<w:spacing '
                  l_substyle
                  '/>'
                  into e_xml respecting blanks.
    endif.

    clear l_substyle.
    if not is_style-leftindent is initial and is_style-leftindent co '0123456789 '.
      data l_indent type string.
      l_indent = is_style-leftindent.
      condense l_indent no-gaps.
      concatenate ' w:left="'
                  l_indent
                  '"'
                  into l_substyle respecting blanks.
    endif.

    if not is_style-rightindent is initial and is_style-rightindent co '0123456789 '.
      l_indent = is_style-rightindent.
      condense l_indent no-gaps.
      concatenate l_substyle
                  ' w:right="'
                  l_indent
                  '"'
                  into l_substyle respecting blanks.
    endif.

    if not is_style-firstindent is initial and is_style-firstindent co '-0123456789 '.
      if is_style-firstindent < 0.
        l_indent = - is_style-firstindent.
        condense l_indent no-gaps.
        concatenate l_substyle
                    ' w:hanging="'
                    l_indent
                    '"'
                    into l_substyle respecting blanks.
      else.
        l_indent = is_style-firstindent.
        condense l_indent no-gaps.
        concatenate l_substyle
                    ' w:firstLine="'
                    l_indent
                    '"'
                    into l_substyle respecting blanks.
      endif.
    endif.

    if not l_substyle is initial.
      concatenate e_xml
                  '<w:ind '
                  l_substyle
                  '/>'
                  into e_xml respecting blanks.
    endif.

* Borders
    clear l_substyle.
    if not is_style-border_left-style is initial
    and not is_style-border_left-width is initial.
      l_size = is_style-border_left-width.
      condense l_size no-gaps.
      l_space = is_style-border_left-space.
      condense l_space no-gaps.
      concatenate l_substyle
                  '<w:left w:val="'
                  is_style-border_left-style
                  '" w:sz="'
                  l_size
                  '" w:space="'
                  l_space
                  '" w:color="'
                  is_style-border_left-color
                  '"/>'
                  into l_substyle respecting blanks.
    endif.

    if not is_style-border_top-style is initial
    and not is_style-border_top-width is initial.
      l_size = is_style-border_top-width.
      condense l_size no-gaps.
      l_space = is_style-border_top-space.
      condense l_space no-gaps.
      concatenate l_substyle
                  '<w:top w:val="'
                  is_style-border_top-style
                  '" w:sz="'
                  l_size
                  '" w:space="'
                  l_space
                  '" w:color="'
                  is_style-border_top-color
                  '"/>'
                  into l_substyle respecting blanks.
    endif.

    if not is_style-border_right-style is initial
    and not is_style-border_right-width is initial.
      l_size = is_style-border_right-width.
      condense l_size no-gaps.
      l_space = is_style-border_right-space.
      condense l_space no-gaps.
      concatenate l_substyle
                  '<w:right w:val="'
                  is_style-border_right-style
                  '" w:sz="'
                  l_size
                  '" w:space="'
                  l_space
                  '" w:color="'
                  is_style-border_right-color
                  '"/>'
                  into l_substyle respecting blanks.
    endif.

    if not is_style-border_bottom-style is initial and
       not is_style-border_bottom-width is initial.

      l_size = is_style-border_bottom-width.
      condense l_size no-gaps.

      l_space = is_style-border_bottom-space.
      condense l_space no-gaps.

      concatenate l_substyle
                  '<w:bottom w:val="'
                  is_style-border_bottom-style
                  '" w:sz="'
                  l_size
                  '" w:space="'
                  l_space
                  '" w:color="'
                  is_style-border_bottom-color
                  '"/>'
                  into l_substyle respecting blanks.

    endif.

    " List
    if is_style-numbering is not initial.

      if is_style-numbering-level is initial.
        data l_numbering_level type string.
        l_numbering_level = '0'.
      else.
        l_numbering_level = zcl_abap_static=>condense( is_style-numbering-level ).
      endif.

      l_substyle = l_substyle && '<w:numPr>'.
      l_substyle = l_substyle && '<w:ilvl w:val="' && l_numbering_level && '"/>'.
      l_substyle = l_substyle && '<w:numId w:val="' && zcl_abap_static=>condense( is_style-numbering-id ) && '"/>'.
      l_substyle = l_substyle && '</w:numPr>'.

    endif.

    if not l_substyle is initial.
      concatenate e_xml
                  '<w:pBdr>'
                  l_substyle
                  '</w:pBdr>'
                  into e_xml respecting blanks.
    endif.

* Add section info if required
    if i_section is not initial.
      concatenate e_xml
                  i_section
                  into e_xml.
    endif.

    if e_xml is not initial.
      e_xml = '<w:pPr>' && e_xml && '</w:pPr>'.
    endif.

  endmethod.                    "_build_paragraph_style


  method build_table.

    data : ls_content    type ref to data,
           l_type(1)     type c,                            "#EC NEEDED
           l_lines       type i,
           l_cols        type i,
           l_col         type i,
           l_col_inc     type i,
           l_style_table type i,
           l_merge       type string,
           l_string      type string,
           l_style       type string,
           l_stylep      type string.
    field-symbols <field> type any.
    field-symbols <ls_field_style> type zcl_word_static=>ty_table_style_field.
    field-symbols <line> type any.

    create data ls_content like line of it_data.
    assign ls_content->* to <line>.
    if sy-subrc ne 0.
      return.
    endif.

    " count number of lines and columns of the table
    describe table it_data lines l_lines.
    if l_lines = 0.
      return.
    endif.
    describe field <line> type l_type components l_cols.
    if l_cols = 0.
      return.
    endif.

    " search if data table is simple or have style infos for each field
    assign component 1 of structure <line> to <field>.
    if sy-subrc ne 0.
      return.
    endif.
    describe field <field> type l_type components l_style_table.

    " Write table properties
    clear e_xml.
    concatenate '<w:tbl>'
                '<w:tblPr>'
                into e_xml.

    " Styled table : define style
    if i_style is not initial.

      concatenate e_xml
                  '<w:tblStyle w:val="'
                  i_style
                  '"/>'
                  into e_xml.

      " If defined, overwrite style table layout (no effect without table style defined)
      clear l_style.
      if is_style is not initial.

        data l_tblwidth type string.
        l_tblwidth = is_style-firstrow.

        condense l_tblwidth no-gaps.
        concatenate l_style
                    ' w:firstRow="'
                    l_tblwidth
                    '"'
                    into l_style respecting blanks.

        l_tblwidth = is_style-firstcol.
        condense l_tblwidth no-gaps.
        concatenate l_style
                    ' w:firstColumn="'
                    l_tblwidth
                    '"'
                    into l_style respecting blanks.

        l_tblwidth = is_style-nozebra.
        condense l_tblwidth no-gaps.
        concatenate l_style
                    ' w:noHBand="'
                    l_tblwidth
                    '"'
                    into l_style respecting blanks.

        l_tblwidth = is_style-novband.
        condense l_tblwidth no-gaps.
        concatenate l_style
                    ' w:noVBand="'
                    l_tblwidth
                    '"'
                    into l_style respecting blanks.

        l_tblwidth = is_style-lastrow.
        condense l_tblwidth no-gaps.
        concatenate l_style
                    ' w:lastRow="'
                    l_tblwidth
                    '"'
                    into l_style respecting blanks.

        l_tblwidth = is_style-lastcol.
        condense l_tblwidth no-gaps.
        concatenate l_style
                    ' w:lastColumn="'
                    l_tblwidth
                    '"'
                    into l_style respecting blanks.

        concatenate e_xml
                    '<w:tblLook'
                    l_style
                    '/>'
                    into e_xml respecting blanks.
        clear l_style.

      endif.

    endif.

    " Default not styled table : add border
    if i_style is initial.

      if i_border = zcl_word_static=>c_true.
        concatenate e_xml
                    '<w:tblBorders>'
                    '<w:top w:color="auto" w:space="0" w:sz="4" w:val="single"/>'
                    '<w:left w:color="auto" w:space="0" w:sz="4" w:val="single"/>'
                    '<w:bottom w:color="auto" w:space="0" w:sz="4" w:val="single"/>'
                    '<w:right w:color="auto" w:space="0" w:sz="4" w:val="single"/>'
                    '<w:insideH w:color="auto" w:space="0" w:sz="4" w:val="single"/>'
                    '<w:insideV w:color="auto" w:space="0" w:sz="4" w:val="single"/>'
                    '</w:tblBorders>'
                    into e_xml.
      endif.

    endif.

    " Define table width
    if i_width is initial.

      " If no table width given, set it to "auto"
      e_xml = e_xml && '<w:tblW w:w="0" w:type="auto"/>'.

    else.

      data l_width type string.
      l_width = i_width.

      e_xml = e_xml && '<w:tblW w:w="' && l_width && '"/>'.

    endif.

    if i_indent is not initial.

      data l_indent type string.
      l_indent = i_indent.

      e_xml = e_xml && '<w:tblInd w:w="' && l_indent && '" w:type="dxa"/>'.

    endif.

    e_xml = e_xml && '</w:tblPr>'.

    " Fill table content
    loop at it_data into <line>.

      e_xml = e_xml && '<w:tr>'.

      l_col = 1.

      do l_cols times.

        do 1 times.

          l_col_inc = 1.

          " Filling the cell
          if l_style_table = 0. "fields are plain text

            e_xml = e_xml && '<w:tc>'.

            assign component l_col of structure <line> to <field>.

            l_string =
              zcl_word_static=>build_text(
                i_text = <field> ).

            e_xml = e_xml && '<w:p>' && l_string && '</w:p>'.

          else. " fields have ty_table_style_field structure, apply styles

            assign component l_col of structure <line> to <ls_field_style>.
            if sy-subrc ne 0. "Occurs when merged cell
              exit. "exit do
            endif.

            if <ls_field_style>-hide eq abap_true.
              exit.
            endif.

            concatenate e_xml
                        '<w:tc>'
                        into e_xml.

            clear : l_style, l_stylep.
            if not <ls_field_style>-bgcolor is initial.
              concatenate l_style
                          '<w:shd w:fill="'
                          <ls_field_style>-bgcolor
                          '"/>'
                          into l_style.
            endif.
            if not <ls_field_style>-valign is initial.
              concatenate l_style
                          '<w:vAlign w:val="'
                          <ls_field_style>-valign
                          '"/>'
                          into l_style.
            endif.

            if <ls_field_style>-merge > 1.

              l_col_inc = <ls_field_style>-merge.

              l_merge = <ls_field_style>-merge.
              condense l_merge no-gaps.

              l_style = l_style && '<w:gridSpan w:val="' && l_merge &&  '"/>'.

            endif.

            if <ls_field_style>-vmerge is not initial.

              case <ls_field_style>-vmerge.
                when 'start'.
                  l_style = l_style && '<w:vMerge w:val="restart"/>'.
                when 'continue'.
                  l_style = l_style && '<w:vMerge/>'.
              endcase.

            endif.

            if <ls_field_style>-width is not initial.
              find '%' in <ls_field_style>-width.
              if sy-subrc eq 0.
                l_style = l_style && '<w:tcW w:w="' && <ls_field_style>-width && '" w:type="pct"/>'.
              else.
                l_style = l_style && '<w:tcW w:w="' && <ls_field_style>-width && '" w:type="dxa"/>'.
              endif.
            endif.

            if <ls_field_style>-direction is not initial.
              case <ls_field_style>-direction.
                when 'UpToDown'.
                  l_style = l_style && '<w:textDirection w:val="btLr"/>'.
              endcase.
            endif.

            if not l_style is initial.
              concatenate '<w:tcPr>' l_style '</w:tcPr>' into l_style.
            endif.

            if <ls_field_style>-xml is not initial.

              concatenate e_xml
                          l_style
                          <ls_field_style>-xml
                          into e_xml.

              find '<w:tbl>' in <ls_field_style>-xml.
              if sy-subrc eq 0.
                e_xml = e_xml && '<w:p></w:p>'.
              endif.

            elseif <ls_field_style>-image_id is not initial.

              data l_xml type string.
              l_xml =
                zcl_word_static=>build_image(
                  i_id     = <ls_field_style>-image_id
                  i_width  = <ls_field_style>-image_width
                  i_height = <ls_field_style>-image_height ).

              l_xml = zcl_word_static=>build_paragraph( l_xml ).

              e_xml = e_xml && l_xml.

            else.

              find 'w:p' in <ls_field_style>-text.
              if sy-subrc eq 0.

                concatenate e_xml
                            l_style
                            <ls_field_style>-text
                            into e_xml.

              else.

                l_string =
                  zcl_word_static=>build_text(
                    i_text   = <ls_field_style>-text
                    i_style  = <ls_field_style>-style
                    is_style = <ls_field_style>-style_effect ).

                if not <ls_field_style>-line_style is initial or
                   not <ls_field_style>-line_style_effect is initial.

                  l_stylep =
                    zcl_word_static=>build_paragraph_style(
                      i_style  = <ls_field_style>-line_style
                      is_style = <ls_field_style>-line_style_effect ).

                endif.

                concatenate e_xml
                            l_style
                            '<w:p>'
                            l_stylep
                            l_string
                            '</w:p>'
                            into e_xml.

              endif.

            endif.

          endif.

          concatenate e_xml '</w:tc>' into e_xml.

        enddo.

        l_col = l_col + l_col_inc.

      enddo.

      e_xml = e_xml && '</w:tr>'.

    endloop.

    concatenate e_xml '</w:tbl>' into e_xml.

  endmethod.                    "write_table


  method build_text.

    data : l_style   type string,
           l_string  type string,
           l_field   type string,
           l_intsize type i,
           l_char6   type c length 6.
    data : lt_find_result type match_result_tab,
           ls_find_result like line of lt_find_result,
           l_off          type i,
           l_len          type i.

* Get font style section
    l_style =
      zcl_word_static=>build_text_style(
        i_style  = i_style
        is_style = is_style ).

* Escape invalid character
    l_string = zcl_word_static=>protect_string( i_text ).

* Replace fields in content
    if l_string cs '##FIELD#'.
* Regex to search all fields to replace
      find all occurrences of regex '##FIELD#([A-Z ])*##' in l_string results lt_find_result.
      sort lt_find_result by offset descending.
* For each result, replace
      loop at lt_find_result into ls_find_result.
        l_off = ls_find_result-offset + 8.
        l_len = ls_find_result-length - 10.
        case l_string+ls_find_result-offset(ls_find_result-length).
          when zcl_word_static=>c_field_pagecount or zcl_word_static=>c_field_pagetotal.
            l_field = l_string+l_off(l_len) && ' \* Arabic'.
          when zcl_word_static=>c_field_filename.
            l_field = l_string+l_off(l_len) && ' \p'.
          when zcl_word_static=>c_field_creationdate or zcl_word_static=>c_field_moddate or zcl_word_static=>c_field_todaydate.
            l_field = l_string+l_off(l_len) && ' \@ &quot;dd/MM/yyyy&quot;'.
          when others.
            l_field = l_string+l_off(l_len).
        endcase.
        concatenate '<w:fldSimple w:instr="'
                    l_field
                    ' \* MERGEFORMAT"/>'
                    into l_field respecting blanks.
        replace l_string+ls_find_result-offset(ls_find_result-length)
                in l_string with l_field.
      endloop.
    endif.

* Replace label anchor by it's value
    if not is_style-label is initial and l_string cs zcl_word_static=>c_label_anchor.

      l_field = protect_label( is_style-label ).

      concatenate '<w:fldSimple w:instr=" SEQ '
                  l_field
                  ' \* ARABIC "/>'
                  into l_field respecting blanks.
      replace zcl_word_static=>c_label_anchor in l_string with l_field.

    endif.

    concatenate '<w:r>'
               l_style
               '<w:t xml:space="preserve">'
               l_string
               '</w:t>'
               '</w:r>'
               into e_xml respecting blanks.

  endmethod.                    "write_text


  method build_text_style.

    if i_style is not initial.
      e_xml = '<w:rStyle w:val="' && i_style && '"/>'.
    endif.

    if is_style is not initial.

      if not is_style-color is initial.
        concatenate e_xml
                    '<w:color w:val="'
                    is_style-color
                    '"/>'
                    into e_xml.
      endif.

      if not is_style-bgcolor is initial.
        concatenate e_xml
                    '<w:shd w:val="clear" w:color="auto" w:fill="'
                    is_style-bgcolor
                    '"/>'
                    into e_xml.
      endif.

      if is_style-bold = zcl_word_static=>c_true.
        concatenate e_xml
                    '<w:b/>'
                    into e_xml.
      endif.

      if is_style-italic = zcl_word_static=>c_true.
        concatenate e_xml
                    '<w:i/>'
                    into e_xml.
      endif.

      if is_style-underline = zcl_word_static=>c_true.
        concatenate e_xml
                    '<w:u w:val="single"/>'
                    into e_xml.
      endif.

      if is_style-strike = zcl_word_static=>c_true.
        concatenate e_xml
                    '<w:strike/>'
                    into e_xml.
      endif.

      if is_style-caps = zcl_word_static=>c_true.
        concatenate e_xml
                    '<w:caps/>'
                    into e_xml.
      endif.

      if is_style-smallcaps = zcl_word_static=>c_true.
        concatenate e_xml
                    '<w:smallCaps/>'
                    into e_xml.
      endif.

      if not is_style-highlight is initial.
        concatenate e_xml
                    '<w:highlight w:val="'
                    is_style-highlight
                    '"/>'
                    into e_xml.
      endif.

      if not is_style-spacing is initial and is_style-spacing co '0123456789 -'.
        if is_style-spacing gt 0.
          data l_string type string.
          l_string = is_style-spacing.
          condense l_string no-gaps.
        else.
          l_string = - is_style-spacing.
          condense l_string no-gaps.
          concatenate '-' l_string into l_string.
        endif.
        concatenate e_xml
                    '<w:spacing w:val="'
                    l_string
                    '"/>'
                    into e_xml.
      endif.

      if is_style-size is not initial.

        data l_intsize type i.
        l_intsize = is_style-size * 2.

        data l_char6(6).
        l_char6 = l_intsize.

        condense l_char6 no-gaps.

        concatenate e_xml
                    '<w:sz w:val="'
                    l_char6
                    '"/>'
                    '<w:szCs w:val="'
                    l_char6
                    '"/>'
                    into e_xml.
      endif.

      if is_style-sup = zcl_word_static=>c_true.
        concatenate e_xml
                    '<w:vertAlign w:val="superscript"/>'
                    into e_xml.
      elseif is_style-sub = zcl_word_static=>c_true.
        concatenate e_xml
                    '<w:vertAlign w:val="subscript"/>'
                    into e_xml.
      endif.

      if not is_style-font is initial.
        concatenate e_xml
                    '<w:rFonts w:ascii="'
                    is_style-font
                    '" w:hAnsi="'
                    is_style-font
                    '"/>'
                    into e_xml.
      endif.

    endif.

    if i_style is not initial.

      concatenate e_xml
                  '<w:rStyle w:val="'
                  i_style
                  '"/>'
                  into e_xml.
    endif.

    if e_xml is not initial.
      e_xml = '<w:rPr>' && e_xml && '</w:rPr>'.
    endif.

  endmethod.                    "build_character_style


  method get_containers_id.

    data lt_result type table of match_result.
    find all occurrences of regex '<w:tag w:val="([^"]*)"/>' in i_xml results lt_result.

    data ls_result like line of lt_result.
    loop at lt_result into ls_result.

      data ls_submatch like line of ls_result-submatches.
      loop at ls_result-submatches into ls_submatch.

        data l_id like line of et_id.
        l_id = i_xml+ls_submatch-offset(ls_submatch-length).

        collect l_id into et_id.

      endloop.

    endloop.

    sort et_id.

  endmethod.


  method get_container_content.

    data ls_result1 type match_result.
    find first occurrence of '<w:tag w:val="' && i_id && '"/>' in i_xml results ls_result1.
    check sy-subrc eq 0.

    data ls_result2 type match_result.
    find first occurrence of '<w:sdtContent>' in i_xml+ls_result1-offset results ls_result2.
    check sy-subrc eq 0.

    data ls_result3 type match_result.
    find first occurrence of '</w:sdtContent>' in i_xml+ls_result1-offset results ls_result3.
    check sy-subrc eq 0.

    data l_offset type i.
    l_offset = ls_result1-offset + ls_result2-offset + ls_result2-length.

    data l_length type i.
    l_length = ls_result1-offset + ls_result3-offset - l_offset.

    e_xml = i_xml+l_offset(l_length).

***  find regex '<w:sdt>.*<w:sdtPr>.*<w:tag w:val="' && i_id && '"\/>.*<\/w:sdtPr>.*<w:sdtContent>(.*)<\/w:sdtContent>.*<\/w:sdt>' in i_xml submatches e_xml.
***  replace regex '(<\/w:sdtContent>.*)' in e_xml with ``.

  endmethod.


  method get_variables.

    data lt_result type match_result_tab.
    find all occurrences of regex '%([0-9A-Za-z_]+)%' in i_xml results lt_result.

    data ls_result like line of lt_result.
    loop at lt_result into ls_result.

      data ls_submatch like line of ls_result-submatches.
      loop at ls_result-submatches into ls_submatch.

        data l_value type string.
        l_value = i_xml+ls_submatch-offset(ls_submatch-length).

        collect l_value into et_list.

      endloop.

    endloop.

    sort et_list.

  endmethod.


  method protect_label.

    out = in.

    translate out using ' _'.

  endmethod.                    "protect_label


  method protect_string.

    out = in.

    replace all occurrences of '&' in out with '&amp;'.
    replace all occurrences of '<' in out with '&lt;'.
    replace all occurrences of '>' in out with '&gt;'.
    replace all occurrences of '"' in out with '&quot;'.

  endmethod.                    "_protect_string


  method set_container_content.

    data ls_result1 type match_result.
    find first occurrence of '<w:tag w:val="' && i_id && '"/>' in c_xml results ls_result1.
    check sy-subrc eq 0.

    data ls_result2 type match_result.
    find first occurrence of '<w:sdtContent>' in c_xml+ls_result1-offset results ls_result2.
    check sy-subrc eq 0.

    data ls_result3 type match_result.
    find first occurrence of '</w:sdtContent>' in c_xml+ls_result1-offset results ls_result3.
    check sy-subrc eq 0.

    data l_length type i.
    l_length = ls_result1-offset + ls_result2-offset + ls_result2-length.

    data l_offset type i.
    l_offset = ls_result1-offset + ls_result3-offset.

    c_xml = c_xml(l_length) && i_xml && c_xml+l_offset.

  endmethod.
ENDCLASS.
