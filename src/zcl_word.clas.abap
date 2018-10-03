*----------------------------------------------------------------------*
*       CLASS ZCL_WORD DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
class zcl_word definition
  public
  create public .

  public section.

*****************************************************************************
* Method constructor
* Constructor method create the word objet and initialize some general data
* - tpl : Empty document to use as a template (docx, dotx, docm, dotm)
*         You could use template in SAPWR instead
*         Use this syntax : SAPWR:<objname>
* - keep_tpl_content:keep the template content as begin of document content
*         (boolean)
*         Default : false
*****************************************************************************
    methods constructor
      importing
        !i_data type xstring optional
      raising
        zcx_generic .
*****************************************************************************
* Method set_title        #### Obsolete method, please use set_properties ###
* Define title for the document
* - title : title of the document
*****************************************************************************
    methods set_title       "obsolete, kept for compatibility
      importing
        !i_text type simple optional
      raising
        zcx_generic .
    class zcl_word_static definition load .
    methods set_header_footer
      importing
        !type               type string default zcl_word_static=>c_type_header
        !textline           type string
        !usenow_default     type i default zcl_word_static=>c_true
        !usenow_first       type i default zcl_word_static=>c_true
        !style              type string optional
        !style_effect       type zcl_word_static=>ty_character_style_effect optional
        !line_style         type string optional
        !line_style_effect  type zcl_word_static=>ty_paragraph_style_effect optional
      exporting
        !id                 type string
        !invalid_style      type i
        !invalid_line_style type i
      raising
        zcx_generic .
    methods set_header_footer_direct
      importing
        !i_header       type string optional
        !i_header_first type string optional
        !i_footer       type string optional
        !i_footer_first type string optional .
    methods set_document_properties
      importing
        !i_title        type simple optional
        !i_author       type simple optional
        !i_description  type simple optional
        !i_object       type simple optional
        !i_category     type simple optional
        !i_keywords     type simple optional
        !i_status       type simple optional
        !i_creationdate type d optional
        !i_creationtime type t optional
        !i_revision     type i optional
        !i_nospellcheck type i optional
      raising
        zcx_generic .
    methods set_page_properties
      importing
        !i_orientation   type i default zcl_word_static=>c_orient_portrait
        !i_border_left   type zcl_word_static=>ty_border optional
        !i_border_top    type zcl_word_static=>ty_border optional
        !i_border_right  type zcl_word_static=>ty_border optional
        !i_border_bottom type zcl_word_static=>ty_border optional .
    methods insert_text
      importing
        !i_text   type string
        !i_style  type string optional
        !is_style type zcl_word_static=>ty_character_style_effect optional .
    methods insert_symbol
      importing
        !i_symbol type simple .
    methods insert_note
      importing
        !text               type string
        !type               type i default zcl_word_static=>c_notetype_foot
        !style              type string optional
        !style_effect       type zcl_word_static=>ty_character_style_effect optional
        !line_style         type string optional
        !line_style_effect  type zcl_word_static=>ty_paragraph_style_effect optional
        !link_style         type string optional
        !link_style_effect  type zcl_word_static=>ty_character_style_effect optional
      exporting
        !invalid_style      type i
        !invalid_link_style type i
        !invalid_line_style type i
      raising
        zcx_generic .
    methods insert_break
      importing
        !i_breaktype type i default zcl_word_static=>c_breaktype_line
          preferred parameter i_breaktype .
    methods insert_paragraph
      importing
        !i_style  type simple optional
        !is_style type zcl_word_static=>ty_paragraph_style_effect optional .
    methods insert_table
      importing
        !it_data  type standard table
        !i_style  type simple optional
        !is_style type zcl_word_static=>ty_style_table optional
        !i_border type i default zcl_word_static=>c_true
        !i_width  type simple optional .
    methods insert_newpage .
    methods insert_toc
      importing
        !default type string optional
        !label   type string optional .
    methods insert_comment
      importing
        !text               type string
        !style              type string optional
        !style_effect       type zcl_word_static=>ty_character_style_effect optional
        !line_style         type string optional
        !line_style_effect  type zcl_word_static=>ty_paragraph_style_effect optional
        !head_style         type string optional
        !head_style_effect  type zcl_word_static=>ty_character_style_effect optional
        !datum              type d default sy-datum
        !uzeit              type t default sy-uzeit
        !author             type string optional
        !initials           type string optional
      exporting
        !invalid_style      type i
        !invalid_link_style type i
        !invalid_line_style type i
      raising
        zcx_generic .
    methods insert_image
      importing
        !i_id           type simple
        !i_zoom         type f optional
        !i_style        type simple optional
        !is_style       type zcl_word_static=>ty_paragraph_style_effect optional
        !i_pattern      type simple optional
        !i_container_id type simple optional .
    methods add_character_style
      importing
        !i_id     type string
        !is_style type zcl_word_static=>ty_character_style_effect
        !i_id_ref type string optional
      raising
        zcx_generic .
    methods add_paragraph_style
      importing
        !i_id          type string
        !is_style      type zcl_word_static=>ty_character_style_effect
        !is_line_style type zcl_word_static=>ty_paragraph_style_effect
        !i_id_ref      type string optional
      raising
        zcx_generic .
    methods add_note
      importing
        !text              type string
        !type              type i
        !style             type string optional
        !style_effect      type zcl_word_static=>ty_character_style_effect optional
        !line_style        type string optional
        !line_style_effect type zcl_word_static=>ty_paragraph_style_effect optional
        !link_style        type string optional
        !link_style_effect type zcl_word_static=>ty_character_style_effect optional
      returning
        value(id)          type string
      raising
        zcx_generic .
    methods add_numbering
      returning
        value(e_id) type string .
    methods add_image
      importing
        !i_name     type simple
        !i_data     type xstring
      returning
        value(e_id) type string
      raising
        zcx_generic .
*****************************************************************************
* Method insert_virtual_field
* Insert a temporary anchor in the document
* The objective of this anchor is to write something here later
* It is usefull when you dont have the content when you write this part,
* but you will have before end of the document
* Use replace_virtual_field to replace the anchor by a viewable content
* - field : Name of the anchor to insert
*****************************************************************************
    methods insert_virtual_field
      importing
        !i_name type simple .
*****************************************************************************
* Method replace_virtual_field
* replace an anchor created with insert_virtual_field
* - field : Name of the anchor to replace
* - value : text content to insert
* - style : Character style name
*           Use only "character" style here.
*           Be carrefull, you must use "internal" name of the style
*           Generaly, it's the same as external without space
*           Default : document default
* - style_effect:Detailed style effect to apply to text fragment
*           Check documentation of class type zcl_word_static=>TY_character_style_effect
* - invalid_style:Character style name given in importing parameter does not
*           exist in document (boolean)
*****************************************************************************
    methods replace_virtual_field
      importing
        !i_name   type simple
        !i_value  type simple
        !i_style  type simple optional
        !is_style type zcl_word_static=>ty_character_style_effect optional .
*****************************************************************************
* Method insert_custom_field
* Insert a custom field link in the document.
* It is not required that the field exist
* - field : Name of the custom field to insert
*****************************************************************************
    methods insert_custom_field
      importing
        !i_name type simple .
    methods set_custom_field
      importing
        !i_name  type simple
        !i_value type simple
      raising
        zcx_generic .
    methods get_image_data
      importing
        !i_id         type simple
      returning
        value(e_data) type xstring .
*****************************************************************************
* Method draw_init
* Initialize draw canvas (the paperboard). All further drawed objects are
* included in this canvas. No one can overflow his size.
* - left    : Left position to start the canvas (in pt)
* - top     : Top position to start the canvas (in pt)
* - width   : Length of the canvas
* - height  : Height of the canvas
* - bgcolor : Background color of the canvas
*             You can use the predefined font color constants
*             or specify any rgb hexa color code
*             Default : transparent
* - bdcolor : Border color of the canvas
*             You can use the predefined font color constants
*             or specify any rgb hexa color code
*             You must specify both bdcolor & bdwidth to have effect applied
*             Default : transparent
* - bdwidth : Border width of the canvas in pt
*             You must specify both bdcolor & bdwidth to have effect applied
*             Default : none
*****************************************************************************
    methods draw_init
      importing
        !left    type i
        !top     type i
        !width   type i
        !height  type i
        !bgcolor type string optional
        !bdcolor type string optional
        !bdwidth type f optional .
*****************************************************************************
* Method draw
* Draw an object in a canvas (you must create it before with draw_init)
* - object  : Type of object to draw
* - left    : Left position to start the canvas (in pt)
* - top     : Top position to start the canvas (in pt)
* - width   : Length of the canvas
* - height  : Height of the canvas
* - url     : In case of image, url of the picture to load
*             You could use image in SAPWR instead
*             Use this syntax : SAPWR:<objname>
* - bgcolor : Background color of the object
*             You can use the predefined font color constants
*             or specify any rgb hexa color code
*             Default : transparent for image, white for other objects
* - bdcolor : Border color of the object
*             You can use the predefined font color constants
*             or specify any rgb hexa color code
*             You must specify both bdcolor & bdwidth to have effect applied
*             Default : transparent for image, black for other objects
* - bdwidth : Border width of the object in pt
*             You must specify both bdcolor & bdwidth to have effect applied
*             Default : none for image, 1px for other objects
* - invalid_image:Image to insert does not exist (boolean)
* - ID (input): You could give ID of an existing image in document instead of
*             url
* - ID (output): ID of the image in document
*****************************************************************************
    methods draw
      importing
        !object   type i
        !left     type i default 0
        !top      type i default 0
        !width    type i default 0
        !height   type i default 0
        !image_id type simple optional
        !bgcolor  type string optional
        !bdcolor  type string optional
        !bdwidth  type f optional .
*****************************************************************************
* Method draw_finalize
* Write all the canvas object in document
* Call this method once you have finished your draw all yours objects.
*****************************************************************************
    methods draw_finalize .
    methods get_docxml
      returning
        value(e_xml) type string .
    methods set_docxml
      importing
        !i_xml type simple .
    methods replace
      importing
        !i_from type simple
        !i_to   type simple .
    methods add_container
      importing
        !i_id type simple .
    methods set_container_content
      importing
        !i_id  type simple
        !i_xml type simple .
    methods get_container_content
      importing
        !i_id        type simple
      returning
        value(e_xml) type string .
    methods get_data
      returning
        value(e_data) type xstring .
  private section.

    data r_zip type ref to cl_abap_zip .
    class zcl_word_static definition load .
    data t_list_style type zcl_word_static=>ty_list_style_table .
    data t_list_object type zcl_word_static=>ty_list_object_table .
    data:
      begin of s_section,
        landscape     type i,
        continuous    type i,
        header_first  type string,
        header        type string,
        footer_first  type string,
        footer        type string,
        border_left   type zcl_word_static=>ty_border,
        border_top    type zcl_word_static=>ty_border,
        border_right  type zcl_word_static=>ty_border,
        border_bottom type zcl_word_static=>ty_border,
      end of s_section .
    data ns_xml type string .
    data doc_xml type string .
    data section_xml type string .
    data tpl_section_xml type string .
    data numbering_xml type string .
    data temp_xml type string .

    methods build_section .
    methods get_file
      importing
        !i_path       type string
      returning
        value(e_text) type string
      raising
        zcx_generic .
    methods set_file
      importing
        !i_path type simple
        !i_text type simple .
    methods delete_file
      importing
        !i_path type simple .
ENDCLASS.



CLASS ZCL_WORD IMPLEMENTATION.


  method add_character_style.

    " Build character style
    data l_style type string.
    l_style =
      zcl_word_static=>build_text_style(
        is_style = is_style ).

    if i_id_ref is not initial.
      data l_string type string.
      concatenate '<w:basedOn w:val="'
                  i_id_ref
                  '"/>'
                  into l_string.
    endif.

    concatenate '<w:style w:type="'
                zcl_word_static=>c_type_character
                '" w:customStyle="1" w:styleId="'
                i_id
                '">'
                '<w:name w:val="'
                i_id
                '"/>'
                l_string
                l_style
                '</w:style>'
                '</w:styles>'
                into l_style.

    " Get styles file
    data l_file type string.
    l_file = get_file( 'word/styles.xml' ).

    " Update style file content
    replace '</w:styles>' in l_file with l_style.

    " Update zipped style file
    set_file(
      i_path = 'word/styles.xml'
      i_text = l_file ).

    " Update style list
    data ls_list_style like line of t_list_style.
    ls_list_style-type = zcl_word_static=>c_type_character.
    ls_list_style-name = i_id.
    append ls_list_style to t_list_style.

  endmethod.                    "create_character_style


  method add_container.

    data l_xml type string.
    l_xml = zcl_word_static=>build_container( i_id ).

    doc_xml = doc_xml && l_xml.

  endmethod.


  method add_image.

    check i_name is not initial.
    check i_data is not initial.

    data l_extension type string.
    l_extension = zcl_text_static=>lower_case( zcl_file_static=>get_extension( i_name ) ).

    " Search available image name
    do.

      data l_filename type string.
      l_filename = 'word/media/image' && sy-index && '.' && l_extension.

      read table r_zip->files transporting no fields
        with key
          name = l_filename.
      if sy-subrc ne 0.
        exit.
      endif.

    enddo.

    " Add image in ZIP
    call method r_zip->add
      exporting
        name    = l_filename
        content = i_data.

    " Get file extension list
    data l_file type string.
    l_file = get_file( '[Content_Types].xml' ).

    " Search if file extension exist
    data l_string type string.
    concatenate 'extension="' l_extension '"' into l_string.

    find first occurrence of l_string in l_file ignoring case.
    if sy-subrc ne 0.

      " If extension is not yet declared, it's time !
      case l_extension.
        when 'jpg'.
          replace '</Types>' with '<Default ContentType="image/jpeg" Extension="jpg"/></Types>'
                  into l_file.
        when 'png'.
          replace '</Types>' with '<Default ContentType="image/png" Extension="png"/></Types>'
                  into l_file.
        when 'gif'.
          replace '</Types>' with '<Default ContentType="image/gif" Extension="gif"/></Types>'
                  into l_file.
      endcase.

      " Update file extension list
      set_file(
        i_path = '[Content_Types].xml'
        i_text = l_file ).

    endif.

    " Get relation file
    l_file = get_file( 'word/_rels/document.xml.rels' ).

    " Create Image ID
    do.

      e_id = 'rId' && zcl_abap_static=>value2text( sy-index ).

      l_string = 'Id="' && e_id && '"'.

      find first occurrence of l_string in l_file ignoring case.
      if sy-subrc ne 0.
        exit. "exit do
      endif.

    enddo.

    " Update object list
    data ls_list_object like line of t_list_object.
    ls_list_object-id   = e_id.
    ls_list_object-type = zcl_word_static=>c_type_image.
    ls_list_object-path = l_filename.
    append ls_list_object to t_list_object.

    " Add relation
    l_filename = l_filename+5.

    concatenate '<Relationship Target="'
                l_filename
                '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Id="'
                e_id
                '"/>'
                '</Relationships>'
                into l_string.

    replace '</Relationships>' with l_string into l_file.

    " Update relation file
    set_file(
      i_path = 'word/_rels/document.xml.rels'
      i_text = l_file ).

  endmethod.                    "_load_image


  method add_note.

    data : l_filename   type string,
           l_string     type string,
           l_file       type string,
           l_link_style type string,
           l_line_style type string,
           l_text       type string,
           l_xmlns      type string,
           l_id         type string.

    " Search if notes file exist
    if type = zcl_word_static=>c_notetype_foot.
      l_filename = 'word/footnotes.xml'.
    elseif type = zcl_word_static=>c_notetype_end.
      l_filename = 'word/endnotes.xml'.
    else.
      return.
    endif.

    read table r_zip->files transporting no fields
      with key
        name = l_filename.

    if sy-subrc = 0.

      " If foot/end notes exists, load the file
      l_file = get_file( l_filename ).

    else.

      " If footnotes doesnt exist, declare it and create it
      " Add footnotes in content_types
      l_file = get_file( '[Content_Types].xml' ).

      if type = zcl_word_static=>c_notetype_foot.
        concatenate '<Override'
                    ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"'
                    ' PartName="/word/footnotes.xml"/></Types>'
                    into l_string respecting blanks.
      elseif type = zcl_word_static=>c_notetype_end.
        concatenate '<Override'
                    ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"'
                    ' PartName="/word/endnotes.xml"/></Types>'
                    into l_string respecting blanks.
      endif.
      replace '</Types>' with l_string
              into l_file.

      set_file(
        i_path = '[Content_Types].xml'
        i_text = l_file ).

      " Add footnotes in relation file
      l_file = get_file( 'word/_rels/document.xml.rels' ).

      " Create footnotes relation ID
      do.
        l_id = 'rId' && sy-index.                           "#EC NOTEXT
        l_string = 'Id="' && l_id && '"'.                   "#EC NOTEXT
        find first occurrence of l_string in l_file ignoring case.
        if sy-subrc ne 0.
          exit. "exit do
        endif.
      enddo.

      " Add relation
      if type = zcl_word_static=>c_notetype_foot.
        concatenate '<Relationship Target="footnotes.xml"'
                    ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"'
                    ' Id="'
                    l_id
                    '"/>'
                    '</Relationships>'
                    into l_string respecting blanks.
      elseif type = zcl_word_static=>c_notetype_end.
        concatenate '<Relationship Target="endnotes.xml"'
                    ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"'
                    ' Id="'
                    l_id
                    '"/>'
                    '</Relationships>'
                    into l_string respecting blanks.
      endif.
      replace '</Relationships>' with l_string into l_file.

      " Update relation file
      set_file(
        i_path = 'word/_rels/document.xml.rels'
        i_text = l_file ).

      " Create notes file
      if type = zcl_word_static=>c_notetype_foot.
        concatenate '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<w:footnotes '
                    ns_xml
                    '>'
                    '<w:footnote w:id="-1" w:type="separator">'
                    '<w:p>'
                    '<w:pPr><w:spacing w:lineRule="auto" w:line="240" w:after="0"/></w:pPr>'
                    '<w:r><w:separator/></w:r>'
                    '</w:p>'
                    '</w:footnote>'
                    '<w:footnote w:id="0" w:type="continuationSeparator">'
                    '<w:p>'
                    '<w:pPr><w:spacing w:lineRule="auto" w:line="240" w:after="0"/></w:pPr>'
                    '<w:r><w:continuationSeparator/></w:r>'
                    '</w:p>'
                    '</w:footnote>'
                    '</w:footnotes>'
                    into l_file respecting blanks.
      elseif type = zcl_word_static=>c_notetype_end.
        concatenate '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<w:endnotes '
                    ns_xml
                    '>'
                    '<w:endnote w:id="-1" w:type="separator">'
                    '<w:p>'
                    '<w:pPr><w:spacing w:lineRule="auto" w:line="240" w:after="0"/></w:pPr>'
                    '<w:r><w:separator/></w:r>'
                    '</w:p>'
                    '</w:endnote>'
                    '<w:endnote w:id="0" w:type="continuationSeparator">'
                    '<w:p>'
                    '<w:pPr><w:spacing w:lineRule="auto" w:line="240" w:after="0"/></w:pPr>'
                    '<w:r><w:continuationSeparator/></w:r>'
                    '</w:p>'
                    '</w:endnote>'
                    '</w:endnotes>'
                    into l_file respecting blanks.
      endif.
    endif.

    " Search available note id
    do.
      id = sy-index.
      condense id no-gaps.
      l_string = 'w:id="' && id && '"'.                     "#EC NOTEXT
      find first occurrence of l_string in l_file ignoring case.
      if sy-subrc ne 0.
        exit. "exit do
      endif.
    enddo.

    " Add blank at start of note
    l_text = text.
    if l_text is initial or l_text(1) ne space.
      concatenate space l_text into l_text respecting blanks.
    endif.

    l_text =
      zcl_word_static=>build_text(
        i_text   = l_text
        i_style  = style
        is_style = style_effect ).

    if not link_style_effect is initial or not link_style is initial.
      l_link_style =
        zcl_word_static=>build_text_style(
          i_style         = link_style
          is_style = link_style_effect ).
    endif.

    if not line_style_effect is initial or not line_style is initial.

      l_line_style =
        zcl_word_static=>build_paragraph_style(
          i_style  = line_style
          is_style = line_style_effect
          i_section = section_xml ).

      clear section_xml.

    endif.

    " Add note
    if type = zcl_word_static=>c_notetype_foot.
      concatenate '<w:footnote w:id="'
                  id
                  '">'
                  '<w:p>'
                  l_line_style
                  '<w:r>'
                  l_link_style
                  '<w:footnoteRef/>'
                  '</w:r>'
                  l_text
                  '</w:p>'
                  '</w:footnote>'
                  '</w:footnotes>'
                  into l_string respecting blanks.
      replace first occurrence of '</w:footnotes>' in l_file with l_string.
    elseif type = zcl_word_static=>c_notetype_end.
      concatenate '<w:endnote w:id="'
                  id
                  '">'
                  '<w:p>'
                  l_line_style
                  '<w:r>'
                  l_link_style
                  '<w:endnoteRef/>'
                  '</w:r>'
                  l_text
                  '</w:p>'
                  '</w:endnote>'
                  '</w:endnotes>'
                  into l_string respecting blanks.
      replace first occurrence of '</w:endnotes>' in l_file with l_string.
    endif.

    " Update footnotes file
    set_file(
      i_path = l_filename
      i_text = l_file ).

  endmethod.                    "_create_footnote


  method add_numbering.

    data lt_abstract_num_id type table of string.
    lt_abstract_num_id =
      zcl_text_static=>get_text_by_pattern(
        i_text    = numbering_xml
        i_pattern = 'w:abstractNumId="([0-9]*)"' ).

    data l_abstract_num_id type i.
    l_abstract_num_id = zcl_abap_static=>get_max( lt_abstract_num_id ) + 1.

    data l_abstract_num_xml type string.
    l_abstract_num_xml = zcl_word_static=>build_abstract_num( l_abstract_num_id ).

    zcl_text_static=>replace_last(
      exporting
        i_from = '</w:abstractNum>'
        i_to   = '</w:abstractNum>' && l_abstract_num_xml
      changing
        c_text = numbering_xml ).

    data lt_num_id type table of string.
    lt_num_id =
      zcl_text_static=>get_text_by_pattern(
        i_text    = numbering_xml
        i_pattern = 'w:numId="([0-9]*)"' ).

    data l_num_id type i.
    l_num_id = zcl_abap_static=>get_max( lt_num_id ) + 1.

    data l_num_xml type string.
    l_num_xml =
      zcl_word_static=>build_num(
      i_abstract_num_id = l_abstract_num_id
      i_id              = l_num_id ).

    zcl_text_static=>replace_last(
      exporting
        i_from = '</w:num>'
        i_to   = '</w:num>' && l_num_xml
      changing
        c_text = numbering_xml ).

    e_id = zcl_abap_static=>value2text( l_num_id ).

  endmethod.


  method add_paragraph_style.

    " Build character style
    data l_style type string.
    l_style =
      zcl_word_static=>build_text_style(
        is_style = is_style ).

    " Build paragraph style
    data l_stylepr type string.
    l_stylepr =
      zcl_word_static=>build_paragraph_style(
        is_style = is_line_style
        i_section = section_xml ).

    clear section_xml.

    if i_id_ref is not initial.
      data l_string type string.
      concatenate '<w:basedOn w:val="'
                  i_id_ref
                  '"/>'
                  into l_string.
    endif.

    concatenate '<w:style w:type="'
                zcl_word_static=>c_type_paragraph
                '" w:customStyle="1" w:styleId="'
                i_id
                '">'
                '<w:name w:val="'
                i_id
                '"/>'
                l_string
                l_stylepr
                l_style
                '</w:style>'
                '</w:styles>'
                into l_style.

    " Get styles file
    data l_file type string.
    l_file = get_file( 'word/styles.xml' ).

    " Update style file content
    replace '</w:styles>' in l_file with l_style.

    " Update zipped style file
    set_file(
      i_path = 'word/styles.xml'
      i_text = l_file ).

    " Update style list
    data ls_list_style like line of t_list_style.
    ls_list_style-type = zcl_word_static=>c_type_paragraph.
    ls_list_style-name = i_id.
    append ls_list_style to t_list_style.

  endmethod.                    "create_paragraph_style


  method build_section.

    clear section_xml.

    " In case of keep template content, first document section is the last
    " template section
    if not tpl_section_xml is initial.
      section_xml = tpl_section_xml.
      clear tpl_section_xml.
      clear s_section.
      return.
    endif.

    " Define header/footer
    if not s_section-header is initial.
      concatenate section_xml
                  '<w:headerReference w:type="default" r:id="'
                  s_section-header
                  '"/>'
                  into section_xml.
    endif.

    if not s_section-footer is initial.
      concatenate section_xml
                  '<w:footerReference w:type="default" r:id="'
                  s_section-footer
                  '"/>'
                  into section_xml.
    endif.

    if not s_section-header_first is initial.
      concatenate section_xml
                  '<w:headerReference w:type="first" r:id="'
                  s_section-header_first
                  '"/>'
                  into section_xml.
    endif.

    if not s_section-footer_first is initial.
      concatenate section_xml
                  '<w:footerReference w:type="first" r:id="'
                  s_section-footer_first
                  '"/>'
                  into section_xml.
    endif.

    " Define page orientation
    if s_section-landscape = zcl_word_static=>c_true.
      concatenate section_xml
                  '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
                  into section_xml.
    else.
      concatenate section_xml
                  '<w:pgSz w:w="11906" w:h="16838"/>'
                  into section_xml.
    endif.

    " Border ?
    data l_substyle type string.
    if not s_section-border_left-style is initial and
       not s_section-border_left-width is initial.

      data l_size type string.
      l_size = s_section-border_left-width.
      condense l_size no-gaps.

      data l_space type string.
      l_space = s_section-border_left-space.
      condense l_space no-gaps.

      concatenate l_substyle
                  '<w:left w:val="'
                  s_section-border_left-style
                  '" w:sz="'
                  l_size
                  '" w:space="'
                  l_space
                  '" w:color="'
                  s_section-border_left-color
                  '"/>'
                  into l_substyle respecting blanks.

    endif.

    if not s_section-border_top-style is initial and
       not s_section-border_top-width is initial.

      l_size = s_section-border_top-width.
      condense l_size no-gaps.

      l_space = s_section-border_top-space.
      condense l_space no-gaps.

      concatenate l_substyle
                  '<w:top w:val="'
                  s_section-border_top-style
                  '" w:sz="'
                  l_size
                  '" w:space="'
                  l_space
                  '" w:color="'
                  s_section-border_top-color
                  '"/>'
                  into l_substyle respecting blanks.
    endif.

    if not s_section-border_right-style is initial and
       not s_section-border_right-width is initial.

      l_size = s_section-border_right-width.
      condense l_size no-gaps.

      l_space = s_section-border_right-space.
      condense l_space no-gaps.

      concatenate l_substyle
                  '<w:right w:val="'
                  s_section-border_right-style
                  '" w:sz="'
                  l_size
                  '" w:space="'
                  l_space
                  '" w:color="'
                  s_section-border_right-color
                  '"/>'
                  into l_substyle respecting blanks.

    endif.

    if not s_section-border_bottom-style is initial and
       not s_section-border_bottom-width is initial.

      l_size = s_section-border_bottom-width.
      condense l_size no-gaps.

      l_space = s_section-border_bottom-space.
      condense l_space no-gaps.

      concatenate l_substyle
                  '<w:bottom w:val="'
                  s_section-border_bottom-style
                  '" w:sz="'
                  l_size
                  '" w:space="'
                  l_space
                  '" w:color="'
                  s_section-border_bottom-color
                  '"/>'
                  into l_substyle respecting blanks.

    endif.

    if not l_substyle is initial.
      concatenate section_xml
                  '<w:pgBorders w:offsetFrom="page">'
                  l_substyle
                  '</w:pgBorders>'
                  into section_xml respecting blanks.
    endif.

    " Default section values / Standard page
    concatenate section_xml
                '<w:cols w:space="708"/>'
                '<w:docGrid w:linePitch="360"/>'
                '<w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>'
                into section_xml.

    if s_section-continuous = zcl_word_static=>c_true.
      concatenate section_xml
                  '<w:type w:val="continuous"/>'
                  into section_xml.
    endif.

    if not s_section-header_first is initial or
       not s_section-footer_first is initial.
      concatenate section_xml
                  '<w:titlePg/>'
                  into section_xml.
    endif.

    concatenate '<w:sectPr>'
                section_xml
                '</w:sectPr>'
                into section_xml.

    clear s_section.

  endmethod.                    "BUILD_SECTION


  method constructor.

    if i_data is initial.

      try.
          data l_data type xstring.
          l_data = cl_docx_form=>create_form(  ).
        catch cx_openxml_not_allowed
              cx_openxml_not_found
              cx_openxml_format
              cx_docx_form_not_unicode.
          zcx_generic=>raise(
            i_text = 'Cannot create empty doc, please use template' ).
      endtry.

    else.
      l_data = i_data.
    endif.

    " Load docx into zip object
    create object r_zip.
    r_zip->load(
      exporting
        zip             = l_data
      exceptions
        zip_parse_error = 1
        others          = 2 ).
    if sy-subrc ne 0.
      zcx_generic=>raise( ).
    endif.

    " Keep actual content
    data l_file type string.
    l_file = get_file( 'word/document.xml' ).

    find regex '<w:document (.*)<w:body>' in l_file submatches ns_xml.

    find regex '<w:body>(.*)</w:body>' in l_file submatches doc_xml.

    find all occurrences of '<w:sectPr' in doc_xml
      match offset sy-fdpos ignoring case.
    if sy-subrc = 0.
      tpl_section_xml = doc_xml+sy-fdpos.
      doc_xml = doc_xml(sy-fdpos).
    endif.

    l_file = get_file( '[Content_Types].xml' ).

    replace all occurrences of 'wordprocessingml.template' in l_file
      with 'wordprocessingml.document'.

    replace all occurrences of 'template.macroEnabledTemplate' in l_file
      with 'document.macroEnabled'.

    set_file(
      i_path = '[Content_Types].xml'
      i_text = l_file ).

    " Set Author, Creation date and Version number properties
    set_document_properties(
      i_creationdate = sy-datlo
      i_creationtime = sy-timlo
      i_revision     = 1 ).

    " Get style file
    l_file = get_file( 'word/styles.xml' ).

    " Scan style file to search all styles
    data lt_find_result type match_result_tab.
    find all occurrences of regex '<w:style ([^>]*)>' in l_file
      results lt_find_result ignoring case.

    data ls_find_result like line of lt_find_result.
    loop at lt_find_result into ls_find_result.

      data ls_list_style type zcl_word_static=>ty_list_style.
      clear ls_list_style.

      data l_string type string.
      find first occurrence of regex 'w:styleId="([^"]*)"' in section
        offset ls_find_result-offset
        length ls_find_result-length
        of l_file
        submatches l_string
        ignoring case.
      if sy-subrc ne 0.
        continue.
      endif.

      ls_list_style-name = l_string.

      find first occurrence of regex 'w:type="(paragraph|character|numbering|table)"' in section
        offset ls_find_result-offset
        length ls_find_result-length
        of l_file
        submatches l_string
        ignoring case.
      if sy-subrc = 0.
        ls_list_style-type = l_string.
      endif.

      append ls_list_style to t_list_style.

    endloop.

    sort t_list_style by type name.

    " Get relation file
    l_file = get_file( 'word/_rels/document.xml.rels' ).

    " Scan relation file to get all objects
    find all occurrences of regex '<Relationship ([^>]*)/>' in l_file
      results lt_find_result ignoring case.

    loop at lt_find_result into ls_find_result.

      data ls_list_object type zcl_word_static=>ty_list_object.
      clear ls_list_object.

      " Search id of object
      find first occurrence of regex 'Id="([^"]*)"'
        in section
        offset ls_find_result-offset
        length ls_find_result-length
        of l_file
        submatches l_string
        ignoring case.
      if sy-subrc ne 0.
        continue.
      endif.

      ls_list_object-id = l_string.

      " Search type of object
      find first occurrence of regex 'Type=".*(footer|header|image)"'
        in section
        offset ls_find_result-offset
        length ls_find_result-length
        of l_file
        submatches l_string
        ignoring case.
      if sy-subrc ne 0.
        continue.
      endif.

      ls_list_object-type = l_string.

      " Search path of file
      find first occurrence of regex 'Target="([^"]*)"'
        in section
        offset ls_find_result-offset
        length ls_find_result-length
        of l_file
        submatches l_string
        ignoring case.
      if sy-subrc ne 0.
        continue.
      endif.

      concatenate 'word/' l_string into ls_list_object-path.

      append ls_list_object to t_list_object.

    endloop.

    sort t_list_object by type id.

    " Numbering
    numbering_xml = get_file( 'word/numbering.xml' ).

  endmethod.                    "constructor


  method delete_file.

    r_zip->delete(
      name = i_path ).

  endmethod.


  method draw.

    data : l_width    type i,
           l_string_x type string,
           l_string_y type string,
           l_string_w type string,
           l_string_h type string,
           l_style    type string,
           l_string   type string,
           l_color    type string.

    l_string_x = l_width = zcl_word_static=>c_basesize * left.
    condense l_string_x no-gaps.

    l_string_y = l_width = zcl_word_static=>c_basesize * top.
    condense l_string_y no-gaps.

    l_string_w = l_width = zcl_word_static=>c_basesize * width.
    condense l_string_w no-gaps.

    l_string_h = l_width = zcl_word_static=>c_basesize * height.
    condense l_string_h no-gaps.

    case object.
      when zcl_word_static=>c_draw_image.

        clear l_style.
        if bgcolor is supplied and not bgcolor is initial.
          concatenate l_style
                      '<a:solidFill>'
                      '<a:srgbClr val="'
                      bgcolor
                      '" />'
                      '</a:solidFill>'
                      into l_style.
        endif.


        if bdcolor is supplied and not bdcolor is initial and
           bdwidth is supplied and not bdwidth is initial.

          l_width = zcl_word_static=>c_basesize * bdwidth.
          l_string = l_width.

          condense l_string no-gaps.

          concatenate l_style
                      '<a:ln w="'
                      l_string
                      '">'
                      '<a:solidFill>'
                      '<a:srgbClr val="'
                      bdcolor
                      '" />'
                      '</a:solidFill>'
                      '</a:ln>'
                      into l_style.
        endif.

        concatenate temp_xml
                    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                    '<pic:nvPicPr>'
                    '<pic:cNvPr name="Image" id="1"/>'
                    '<pic:cNvPicPr>'
                    '<a:picLocks/>'
                    '</pic:cNvPicPr>'
                    '</pic:nvPicPr>'
                    '<pic:blipFill>'
                    '<a:blip r:embed="'
                    image_id
                    '">'
                    '</a:blip>'
                    '<a:stretch><a:fillRect/></a:stretch>'
                    '</pic:blipFill>'
                    '<pic:spPr>'
                    '<a:xfrm>'
                    '<a:off y="'
                    l_string_y
                    '" x="'
                    l_string_x
                    '"/>'
                    '<a:ext cy="'
                    l_string_h
                    '" cx="'
                    l_string_w
                    '"/>'
                    '</a:xfrm>'
                    '<a:prstGeom prst="rect"/>'
                    l_style
                    '</pic:spPr>'
                    '</pic:pic>'
                    into temp_xml.

      when zcl_word_static=>c_draw_rectangle.

        " Default : white rect with 1pt black border
        clear l_style.
        if bgcolor is supplied and not bgcolor is initial.
          l_color = bgcolor.
        else.
          l_color = zcl_word_static=>c_color_white.
        endif.
        concatenate l_style
                    '<a:solidFill>'
                    '<a:srgbClr val="'
                    l_color
                    '" />'
                    '</a:solidFill>'
                    into l_style.

        if bdcolor is supplied and not bdcolor is initial.
          l_color = bdcolor.
        else.
          l_color = zcl_word_static=>c_color_black.
        endif.
        if bdwidth is supplied and not bdwidth is initial.
          l_width = zcl_word_static=>c_basesize * bdwidth.
        else.
          l_width = zcl_word_static=>c_basesize.
        endif.
        l_string = l_width.
        condense l_string no-gaps.

        concatenate l_style
                    '<a:ln w="'
                    l_string
                    '">'
                    '<a:solidFill>'
                    '<a:srgbClr val="'
                    l_color
                    '" />'
                    '</a:solidFill>'
                    '</a:ln>'
                    into l_style.

        concatenate temp_xml
                    '<wps:wsp>'
                    '<wps:cNvPr name="Rectangle" id="1"/>'
                    '<wps:cNvSpPr/>'
                    '<wps:spPr>'
                    '<a:xfrm>'
                    '<a:off y="'
                    l_string_y
                    '" x="'
                    l_string_x
                    '"/>'
                    '<a:ext cy="'
                    l_string_h
                    '" cx="'
                    l_string_w
                    '"/>'
                    '</a:xfrm>'
                    '<a:prstGeom prst="rect"/>'
                    l_style
                    '</wps:spPr>'
                    '<wps:bodyPr />'
                    '</wps:wsp>'
                    into temp_xml respecting blanks.
    endcase.

  endmethod.                    "draw


  method draw_finalize.

    concatenate doc_xml
                '<w:p>'
                temp_xml
                '</wpc:wpc>'
                '</a:graphicData>'
                '</a:graphic>'
                '</wp:anchor>'
                '</w:drawing>'
                '</mc:Choice>'
                '<mc:Fallback><w:t>Canva graphic cannot be loaded</w:t></mc:Fallback>'
                '</mc:AlternateContent>'
                '</w:r>'
                '</w:p>'
                into doc_xml.

    clear temp_xml.

  endmethod.                    "draw_finalize


  method draw_init.
    data : l_style    type string,
           l_width    type i,
           l_string   type string,
           l_string_x type string,
           l_string_y type string,
           l_string_w type string,
           l_string_h type string
           .

    clear temp_xml.
    clear l_style.
    if bgcolor is supplied and not bgcolor is initial.
      concatenate l_style
                  '<wpc:bg>'
                  '<a:solidFill>'
                  '<a:srgbClr val="'
                  bgcolor
                  '" />'
                  '</a:solidFill>'
                  '</wpc:bg>'
                  into l_style.
    endif.

    if bdcolor is supplied and not bdcolor is initial
    and bdwidth is supplied and not bdwidth is initial.
      l_width = zcl_word_static=>c_basesize * bdwidth.
      l_string = l_width.
      condense l_string no-gaps.

      concatenate l_style
                  '<wpc:whole>'
                  '<a:ln w="'
                  l_string
                  '">'
                  '<a:solidFill>'
                  '<a:srgbClr val="'
                  bdcolor
                  '" />'
                  '</a:solidFill>'
                  '</a:ln>'
                  '</wpc:whole>'
                  into l_style.
    endif.

    l_string_x = l_width = zcl_word_static=>c_basesize * left.
    condense l_string_x no-gaps.
    l_string_y = l_width = zcl_word_static=>c_basesize * top.
    condense l_string_y no-gaps.
    l_string_w = l_width = zcl_word_static=>c_basesize * width.
    condense l_string_w no-gaps.
    l_string_h = l_width = zcl_word_static=>c_basesize * height.
    condense l_string_h no-gaps.

    concatenate '<w:r>'
                '<mc:AlternateContent>'
                '<mc:Choice Requires="wpc">'
                '<w:drawing>'
                '<wp:anchor distR="0" distL="0" distB="0" distT="0" allowOverlap="0" layoutInCell="1" locked="0" behindDoc="1" relativeHeight="0" simplePos="0">'
                '<wp:simplePos y="0" x="0"/>'
                '<wp:positionH relativeFrom="column">'
                '<wp:posOffset>'
                l_string_x
                '</wp:posOffset>'
                '</wp:positionH>'
                '<wp:positionV relativeFrom="paragraph">'
                '<wp:posOffset>'
                l_string_y
                '</wp:posOffset>'
                '</wp:positionV>'
                '<wp:extent cy="'
                l_string_h
                '" cx="'
                l_string_w
                '"/>'
                '<wp:wrapNone/>'
                '<wp:docPr name="Draw container" id="3"/>'
                '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas">'
                '<wpc:wpc>'
                l_style
                into temp_xml respecting blanks.
  endmethod.                    "draw_init


  method get_container_content.

    e_xml =
      zcl_word_static=>get_container_content(
        i_id  = i_id
        i_xml = doc_xml ).

  endmethod.


  method get_data.

    " Add final section info
    build_section( ).

    " Add complete xml header
    data l_xml type string.
    l_xml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` &&
      `<w:document ` && ns_xml && `>` &&
      `<w:body>` && doc_xml && section_xml && `</w:body>` &&
      `</w:document>`.

    " Add custom body into docx
    set_file(
      i_path = 'word/document.xml'
      i_text = l_xml ).

    " Numbering
    set_file(
      i_path = 'word/numbering.xml'
      i_text  = numbering_xml ).

    e_data = r_zip->save( ).

  endmethod.                    "get_docx_file


  method get_docxml.

    e_xml = doc_xml.

  endmethod.                    "get_DOC_XML


  method get_file.

    " Get zipped file
    data l_data type xstring.
    r_zip->get(
      exporting
        name    = i_path
      importing
        content = l_data
      exceptions
        others  = 1 ).

    e_text = zcl_convert_static=>xtext2text( l_data ).

  endmethod.                    "get_file


  method get_image_data.

    data ls_list_object like line of t_list_object.
    read table t_list_object into ls_list_object
      with key
        type = zcl_word_static=>c_type_image
        id   = i_id.
    check sy-subrc eq 0.

    data l_data type xstring.
    call method r_zip->get
      exporting
        name    = ls_list_object-path
      importing
        content = e_data
      exceptions
        others  = 1.

  endmethod.                    "get_image_data


  method insert_break.

    case i_breaktype.
      when zcl_word_static=>c_breaktype_line.

        temp_xml = temp_xml && '<w:r><w:br/></w:r>'.

      when zcl_word_static=>c_breaktype_page.

        temp_xml = temp_xml && '<w:br w:type="page"/>'.

      when zcl_word_static=>c_breaktype_section.

        build_section( ).

      when zcl_word_static=>c_breaktype_section_continuous.

        build_section( ).

        s_section-continuous = zcl_word_static=>c_true.

      when others.
        return.

    endcase.

    if i_breaktype ne zcl_word_static=>c_breaktype_line.
      insert_paragraph( ).
    endif.

  endmethod.                    "write_break


  method insert_comment.

    data : l_file               type string,
           l_string             type string,
           l_id                 type string,
           l_text               type string,
           l_xmlns              type string,
           l_head_style         type string,
           l_line_style         type string,
           ls_head_style_effect type zcl_word_static=>ty_character_style_effect,
           l_author             type string,
           l_initials           type string.

    read table r_zip->files transporting no fields
      with key name = 'word/comments.xml'.

    if sy-subrc = 0.
      " If comment file exists, load the file
      l_file = get_file( 'word/comments.xml' ).
    else.
      " If comments file doesnt exist, declare it and create it
      " Add comments in content_types
      l_file = get_file( '[Content_Types].xml' ).

      concatenate '<Override'
                  ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"'
                  ' PartName="/word/comments.xml"/></Types>'
                  into l_string respecting blanks.
      replace '</Types>' with l_string
              into l_file.

      set_file(
        i_path = '[Content_Types].xml'
        i_text = l_file ).

      " Add comments in relation file
      l_file = get_file( 'word/_rels/document.xml.rels' ).

      " Create comments relation ID
      do.
        l_id = 'rId' && sy-index.                           "#EC NOTEXT
        l_string = 'Id="' && l_id && '"'.                   "#EC NOTEXT
        find first occurrence of l_string in l_file ignoring case.
        if sy-subrc ne 0.
          exit. "exit do
        endif.
      enddo.

      " Add relation
      concatenate '<Relationship Target="comments.xml"'
                  ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"'
                  ' Id="'
                  l_id
                  '"/>'
                  '</Relationships>'
                  into l_string respecting blanks.
      replace '</Relationships>' with l_string into l_file.

      " Update relation file
      set_file(
        i_path = 'word/_rels/document.xml.rels'
        i_text = l_file ).

      " Create empty comments file
      concatenate '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                  "cl_abap_char_utilities=>cr_lf
                  cl_abap_char_utilities=>newline
                  '<w:comments '
                  ns_xml
                  '>'
                  '</w:comments>'
                  into l_file respecting blanks.
    endif.

    " Search available comment id
    do.
      "sy-index = sy-index + 4.
      l_id = sy-index.
      condense l_id no-gaps.
      l_string = 'w:id="' && l_id && '"'.                   "#EC NOTEXT
      find first occurrence of l_string in l_file ignoring case.
      if sy-subrc ne 0.
        exit. "exit do
      endif.
    enddo.

    " Add blank at start of note
    l_text = text.
    if l_text is initial or l_text(1) ne space.
      concatenate space l_text into l_text respecting blanks.
    endif.

    l_text =
      zcl_word_static=>build_text(
        i_text   = l_text
        i_style  = style
        is_style = style_effect ).

    " Define default style for comment head
    ls_head_style_effect = head_style_effect.
    if ls_head_style_effect is initial and head_style is initial.
      ls_head_style_effect-bold = zcl_word_static=>c_true.
    endif.

    l_head_style =
      zcl_word_static=>build_text_style(
        i_style  = head_style
        is_style = ls_head_style_effect ).

    if not line_style_effect is initial or not line_style is initial.

      l_line_style =
        zcl_word_static=>build_paragraph_style(
          i_style  = line_style
          is_style = line_style_effect
          i_section = section_xml ).

      clear section_xml.

    endif.

    " Define author property
    if author is initial.
      l_author = author.
    else.
      l_author = author.
    endif.

    " Define initial property
    if initials is initial.
      l_initials = l_author.
    else.
      l_initials = initials.
    endif.

    concatenate '<w:comment w:initials="'
                l_initials
                '"'
                ' w:date="'
                datum(4)
                '-'
                datum+4(2)
                '-'
                datum+6(2)
                'T'
                uzeit(2)
                ':'
                uzeit+2(2)
                ':'
                uzeit+4(2)
                'Z"'
                ' w:author="'
                l_author
                '"'
                ' w:id="'
                l_id
                '">'
                '<w:p>'
                l_line_style
                '<w:r>'
                l_head_style
                '<w:annotationRef/>'
                '</w:r>'
                l_text
                '</w:p>'
                '</w:comment>'
                '</w:comments>'
                into l_string respecting blanks.

    replace '</w:comments>' with l_string into l_file.

    set_file(
      i_path = 'word/comments.xml'
      i_text = l_file ).

    " Finally insert reference to comment in current text fragment
    concatenate temp_xml
                '<w:r><w:commentReference w:id="'
                l_id
                '"/></w:r>'
                into temp_xml.

  endmethod.                    "write_comment


  method insert_custom_field.

    temp_xml =
      temp_xml &&
      `<w:fldSimple w:instr=" DOCPROPERTY ` &&  i_name && ` \* MERGEFORMAT ">` &&
      `<w:r><w:rPr><w:b/></w:rPr><w:t>Please update field</w:t></w:r>` &&
      `</w:fldSimple>`.

  endmethod.                    "insert_custom_field


  method insert_image.

    data l_data type xstring.
    l_data = get_image_data( i_id ).

    check l_data is not initial.

    data l_width type i.
    data l_height type i.
    call method cl_fxs_image_info=>determine_info
      exporting
        iv_data = l_data
      importing
        ev_xres = l_width
        ev_yres = l_height.

    if s_section-landscape = zcl_word_static=>c_true.
      data l_landscape.
      l_landscape = abap_true.
    endif.

    data l_xml type string.
    l_xml =
      zcl_word_static=>build_paragraph(
        zcl_word_static=>build_image(
          i_id        = i_id
          i_width     = l_width
          i_height    = l_height
          i_zoom      = i_zoom
          i_landscape = l_landscape ) ).

    " Insert image in document
    if i_pattern is not initial.

      replace(
        i_from = i_pattern
        i_to   = l_xml ).

    elseif i_container_id is not initial.

      set_container_content(
        i_id  = i_container_id
        i_xml = l_xml ).

    else.

      doc_xml = doc_xml && l_xml.

    endif.

  endmethod.                    "insert_image


  method insert_newpage.

    insert_break( zcl_word_static=>c_breaktype_page ).

  endmethod.                    "write_newpage


  method insert_note.

    data : ls_link_style_effect type zcl_word_static=>ty_character_style_effect,
           ls_line_style_effect type zcl_word_static=>ty_paragraph_style_effect,
           l_style              type string,
           l_string             type string,
           l_id                 type string.

    " Define a default style for link
    ls_link_style_effect = link_style_effect.
    if link_style is initial and ls_link_style_effect is initial.
      ls_link_style_effect-sup = zcl_word_static=>c_true.
    endif.

    " Define a default style for footnote
    ls_line_style_effect = line_style_effect.
    if line_style is initial and ls_line_style_effect is initial.
      ls_line_style_effect-spacing_before = '0'.
      ls_line_style_effect-spacing_after = '0'.
      ls_line_style_effect-interline = 240.
    endif.

    " Create a new footnote
    l_id =
      add_note(
        text               = text
        type               = type
        style              = style
        style_effect       = style_effect
        line_style         = line_style
        line_style_effect  = ls_line_style_effect
        link_style         = link_style
        link_style_effect  = ls_link_style_effect ).

    if l_id is initial.
      return.
    endif.

    " Prepare style for the footnote link
    l_style =
      zcl_word_static=>build_text_style(
        i_style         = link_style
        is_style = ls_link_style_effect ).

    " Now insert note in document
    if type = zcl_word_static=>c_notetype_foot.
      l_string = '<w:footnoteReference w:id="'.
    elseif type = zcl_word_static=>c_notetype_end.
      l_string = '<w:endnoteReference w:id="'.
    endif.

    concatenate temp_xml
                '<w:r>'
                l_style
                l_string
                l_id
                '"/>'
                '</w:r>'
                into temp_xml.

  endmethod.                    "write_footnote


  method insert_paragraph.

    data l_xml type string.
    l_xml =
      zcl_word_static=>build_paragraph(
        i_text    = temp_xml
        i_style   = i_style
        is_style  = is_style
        i_section = section_xml ).

    doc_xml = doc_xml && l_xml.

    clear temp_xml.
    clear section_xml.

  endmethod.                    "write_line


  method insert_symbol.

    data ls_style type zcl_word_static=>ty_character_style_effect.
    ls_style-font = zcl_word_static=>c_font_symbol.

    insert_text(
      i_text   = i_symbol
      is_style = ls_style ).

  endmethod.                    "write_symbol


  method insert_table.

    data l_xml type string.
    l_xml =
      zcl_word_static=>build_table(
        it_data  = it_data
        i_style  = i_style
        is_style = is_style
        i_border = i_border
        i_width  = i_width ).

    doc_xml = doc_xml && l_xml.

  endmethod.                    "write_table


  method insert_text.

    data l_xml type string.
    l_xml =
      zcl_word_static=>build_text(
        i_text   = i_text
        i_style  = i_style
        is_style = is_style ).

    temp_xml = temp_xml && l_xml.

  endmethod.                    "build_text


  method insert_toc.

    data : l_default type string,
           l_content type string.

* Classic TOC
    if label is initial.
      l_default = '--== Table of content - please refresh ==--'.
      l_content = '\o "1-9"'.

* Table of content for specific label (table, figure, ...)
    else.
      concatenate '--== Table of '
                  label
                  ' - please refresh ==--'
                  into l_default respecting blanks.

      l_content = zcl_word_static=>protect_label( label ).

      concatenate '\h \h \c "'
                  l_content
                  '"'
                  into l_content.
    endif.

* If specified, use given default text
    if not default is initial.
      l_default = default.
    endif.

* Write TOC
    concatenate doc_xml
                '<w:p>'
                '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                '<w:r><w:instrText> TOC '
                l_content
                ' </w:instrText></w:r>'
                '</w:p>'

* Add a default text for initial TOC value
                '<w:p>'
                '<w:pPr><w:jc w:val="center"/></w:pPr>'
                '<w:r>'
                '<w:fldChar w:fldCharType="separate"/></w:r>'
                '<w:r>'
                '<w:rPr><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr>'
                '<w:t>'
                l_default
                '</w:t></w:r>'
                '</w:p>'

                '<w:p>'
                '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
                '</w:p>'

                into doc_xml respecting blanks.
  endmethod.                    "write_toc


  method insert_virtual_field.

    temp_xml = temp_xml && '<w:virtual>' && i_name && '</w:virtual>'.

  endmethod.                    "insert_virtual_field


  method replace.

    replace all occurrences of i_from in doc_xml
      with i_to.

  endmethod.                    "replace


  method replace_virtual_field.

    data l_field type string.
    concatenate '<w:virtual>'
                i_name
                '</w:virtual>'
           into l_field.

    data l_value type string.
    l_value =
      zcl_word_static=>build_text(
        i_text   = i_value
        i_style  = i_style
        is_style = is_style ).

    replace l_field in temp_xml with l_value ignoring case.
    if sy-subrc ne 0.
      replace l_field in doc_xml with l_value ignoring case.
    endif.

  endmethod.                    "replace_virtual_field


  method set_container_content.

    zcl_word_static=>set_container_content(
      exporting
        i_id  = i_id
        i_xml = i_xml
      changing
        c_xml = doc_xml ).

  endmethod.


  method set_custom_field. "not managed

    data : l_file   type string,
           l_string type string,
           l_id     type string.

    " If customproperties does not exist, create it
    read table r_zip->files transporting no fields
      with key
        name = 'docProps/custom.xml'.

    if sy-subrc ne 0.

      " Declare new file in relations
      l_file = get_file( '_rels/.rels' ).

      " search available id
      do.
        l_id = 'rId' && sy-index.                           "#EC NOTEXT
        l_string = 'Id="' && l_id && '"'.                   "#EC NOTEXT
        find l_string in l_file ignoring case.
        if sy-subrc ne 0.
          exit. "exit do
        endif.
      enddo.

      concatenate '<Relationship Target="docProps/custom.xml"'
                  ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"'
                  ' Id="'
                  l_id
                  '"/></Relationships>'
                  into l_string respecting blanks.
      replace '</Relationships>' in l_file with l_string.

      set_file(
        i_path = '_rels/.rels'
        i_text = l_file ).

      " Declare new file in content type
      l_file = get_file( '[Content_Types].xml' ).

      replace '</Types>' in l_file
              with '<Override ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml" PartName="/docProps/custom.xml"/></Types>'.

      set_file(
        i_path = '[Content_Types].xml'
        i_text = l_file ).

      concatenate '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
                  '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"'
                  ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
                  '</Properties>'
                  into l_file respecting blanks.

      set_file(
        i_path = 'docProps/custom.xml'
        i_text = l_file ).

    endif.

    " Open custom properties
    l_file = get_file( 'docProps/custom.xml' ).

    " Search available ID
    do.

      if sy-index = 1.
        continue.
      endif.
      l_id = sy-index.
      condense l_id no-gaps.
      l_string = 'pid="' && l_id && '"'.
      find l_string in l_file ignoring case.
      if sy-subrc ne 0.
        exit. "exit do
      endif.

    enddo.

    " Add property
    concatenate '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"'
                ' pid="'
                l_id
                '" name="'
                i_name
                '">'
                '<vt:lpwstr>'
                i_value
                '</vt:lpwstr>'
                '</property>'
                '</Properties>'
                into l_string respecting blanks.
    replace '</Properties>' in l_file with l_string.

    " Update custopproperties file
    set_file(
      i_path = 'docProps/custom.xml'
      i_text = l_file ).

  endmethod.                    "create_custom_field


  method set_document_properties.

    " Get prperties file
    data l_xml type string.
    l_xml = get_file( 'docProps/core.xml' )..

    if i_title is supplied.

      " Replace existing title by new one
      data l_string type string.
      l_string = zcl_word_static=>protect_string( i_title ).

      concatenate '<dc:title>'
                  l_string
                  '</dc:title>'
                  into l_string.

      replace first occurrence of regex '<dc:title>(.*)</dc:title>' in l_xml with l_string.
      if sy-subrc ne 0.
        replace first occurrence of regex '<dc:title\s*/>' in l_xml with l_string.
      endif.

      " If no title property found, add it at end of xml
      if sy-subrc ne 0.

        concatenate l_string
                    '</cp:coreProperties>'
                    into l_string.

        replace '</cp:coreProperties>'
                with l_string
                into l_xml.

      endif.

    endif.

    if i_object is supplied.

      " Replace existing object by new one
      l_string = zcl_word_static=>protect_string( i_object ).

      concatenate '<dc:subject>'
                  l_string
                  '</dc:subject>'
                  into l_string.

      replace first occurrence of regex '<dc:subject>(.*)</dc:subject>' in l_xml with l_string.
      if sy-subrc ne 0.
        replace first occurrence of regex '<dc:subject\s*/>' in l_xml with l_string.
      endif.

      " If no object property found, add it at end of xml
      if sy-subrc ne 0.

        concatenate l_string
                    '</cp:coreProperties>'
                    into l_string.

        replace '</cp:coreProperties>'
                with l_string
                into l_xml.

      endif.

    endif.

    if i_author is supplied.

      " Replace existing author by new one
      l_string =
        '<dc:creator>' &&
        zcl_word_static=>protect_string( i_author ) &&
        '</dc:creator>'.

      replace first occurrence of regex '<dc:creator>(.*)</dc:creator>' in l_xml with l_string.
      if sy-subrc ne 0.
        replace first occurrence of regex '<dc:creator\s*/>' in l_xml with l_string.
      endif.

      " If no author property found, add it at end of xml
      if sy-subrc ne 0.
        concatenate l_string
                    '</cp:coreProperties>'
                    into l_string.
        replace '</cp:coreProperties>'
                with l_string
                into l_xml.
      endif.

      " Set also last modified property
      l_string = zcl_word_static=>protect_string( i_author ).

      concatenate '<cp:lastModifiedBy>'
                  l_string
                  '</cp:lastModifiedBy>'
                  into l_string.

      replace first occurrence of regex '<cp:lastModifiedBy>(.*)</cp:lastModifiedBy>' in l_xml with l_string.
      if sy-subrc ne 0.
        replace first occurrence of regex '<cp:lastModifiedBy\s*/>' in l_xml with l_string.
      endif.

      " If no lastmodified property found, add it at end of xml
      if sy-subrc ne 0.

        concatenate l_string
                    '</cp:coreProperties>'
                    into l_string.

        replace '</cp:coreProperties>'
                with l_string
                into l_xml.

      endif.

    endif.

    if i_creationdate is supplied.

      " Replace existing creation date by new one
      concatenate '<dcterms:created xsi:type="dcterms:W3CDTF">'
                  i_creationdate(4)
                  '-'
                  i_creationdate+4(2)
                  '-'
                  i_creationdate+6(2)
                  'T'
                  i_creationtime(2)
                  ':'
                  i_creationtime+2(2)
                  ':'
                  i_creationtime+2(2)
                  'Z'
                  '</dcterms:created>'
                  into l_string.

      replace first occurrence of regex '<dcterms:created(.*)</dcterms:created>' in l_xml with l_string.
      if sy-subrc ne 0.
        replace first occurrence of regex '<dcterms:created\s*/>' in l_xml with l_string.
      endif.

      " If no creation date property found, add it at end of xml
      if sy-subrc ne 0.

        concatenate l_string
                    '</cp:coreProperties>'
                    into l_string.

        replace '</cp:coreProperties>'
                with l_string
                into l_xml.

      endif.

      " Replace also last modification date by new one
      concatenate '<dcterms:modified xsi:type="dcterms:W3CDTF">'
                  i_creationdate(4)
                  '-'
                  i_creationdate+4(2)
                  '-'
                  i_creationdate+6(2)
                  'T'
                  i_creationtime(2)
                  ':'
                  i_creationtime+2(2)
                  ':'
                  i_creationtime+2(2)
                  'Z'
                  '</dcterms:modified>'
                  into l_string.

      replace first occurrence of regex '<dcterms:modified(.*)</dcterms:modified>' in l_xml with l_string.
      if sy-subrc ne 0.
        replace first occurrence of regex '<dcterms:modified\s*/>' in l_xml with l_string.
      endif.

      " If no modification date property found, add it at end of xml
      if sy-subrc ne 0.

        concatenate l_string
                    '</cp:coreProperties>'
                    into l_string.

        replace '</cp:coreProperties>'
                with l_string
                into l_xml.

      endif.

    endif.

    if i_description is supplied.

      " Replace existing description by new one
      l_string = zcl_word_static=>protect_string( i_description ).

      concatenate '<dc:description>'
                  l_string
                  '</dc:description>'
                  into l_string.

      replace first occurrence of regex '<dc:description>(.*)</dc:description>' in l_xml with l_string.
      if sy-subrc ne 0.
        replace first occurrence of regex '<dc:description\s*/>' in l_xml with l_string.
      endif.

      " If no description property found, add it at end of xml
      if sy-subrc ne 0.

        concatenate l_string
                    '</cp:coreProperties>'
                    into l_string.

        replace '</cp:coreProperties>'
                with l_string
                into l_xml.
      endif.

    endif.

    if i_category is supplied.

      " Replace existing category by new one
      l_string = zcl_word_static=>protect_string( i_category ).

      concatenate '<cp:category>'
                  l_string
                  '</cp:category>'
                  into l_string.

      replace first occurrence of regex '<cp:category>(.*)</cp:category>' in l_xml with l_string.
      if sy-subrc ne 0.
        replace first occurrence of regex '<cp:category\s*/>' in l_xml with l_string.
      endif.

      " If no category property found, add it at end of xml
      if sy-subrc ne 0.

        concatenate l_string
                    '</cp:coreProperties>'
                    into l_string.

        replace '</cp:coreProperties>'
                with l_string
                into l_xml.

      endif.

    endif.

    if i_keywords is supplied.

      " Replace existing keywords by new one
      l_string = zcl_word_static=>protect_string( i_keywords ).

      concatenate '<cp:keywords>'
                  l_string
                  '</cp:keywords>'
                  into l_string.

      replace first occurrence of regex '<cp:keywords>(.*)</cp:keywords>' in l_xml with l_string.
      if sy-subrc ne 0.
        replace first occurrence of regex '<cp:keywords\s*/>' in l_xml with l_string.
      endif.

      " If no keywords property found, add it at end of xml
      if sy-subrc ne 0.

        concatenate l_string
                    '</cp:coreProperties>'
                    into l_string.

        replace '</cp:coreProperties>'
                with l_string
                into l_xml.

      endif.

    endif.

    if i_status is supplied.

      " Replace existing status by new one
      l_string = zcl_word_static=>protect_string( i_status ).

      concatenate '<cp:contentStatus>'
                  l_string
                  '</cp:contentStatus>'
                  into l_string.

      replace first occurrence of regex '<cp:contentStatus>(.*)</cp:contentStatus>' in l_xml with l_string.
      if sy-subrc ne 0.
        replace first occurrence of regex '<cp:contentStatus\s*/>' in l_xml with l_string.
      endif.

      " If no status property found, add it at end of xml
      if sy-subrc ne 0.

        concatenate l_string
                    '</cp:coreProperties>'
                    into l_string.

        replace '</cp:coreProperties>'
                with l_string
                into l_xml.

      endif.

    endif.

    if i_revision is supplied.

      " Replace existing status by new one
      l_string = i_revision.
      condense l_string no-gaps.

      concatenate '<cp:revision>'
                  l_string
                  '</cp:revision>'
                  into l_string.

      replace first occurrence of regex '<cp:revision>(.*)</cp:revision>' in l_xml with l_string.
      if sy-subrc ne 0.
        replace first occurrence of regex '<cp:revision\s*/>' in l_xml with l_string.
      endif.

      " If no revision property found, add it at end of xml
      if sy-subrc ne 0.

        concatenate l_string
                    '</cp:coreProperties>'
                    into l_string.

        replace '</cp:coreProperties>'
                with l_string
                into l_xml.

      endif.

    endif.

    " Update properties file
    set_file(
      i_path = 'docProps/core.xml'
      i_text = l_xml ).

    " Hide spellcheck for this document
    if i_nospellcheck is supplied and
       i_nospellcheck eq zcl_word_static=>c_true.

      l_xml = get_file( 'word/settings.xml' ).

      " File doesnt exist ? create it !
      if l_xml is initial.

        concatenate '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<w:settings '
                    ns_xml
                    '>'
                    '</w:settings>'
                    into l_xml respecting blanks.
      endif.

      find 'hideSpellingErrors' in l_xml.
      if sy-subrc ne 0.
        replace first occurrence of '</w:settings>'
          in l_xml
          with '<w:hideSpellingErrors/></w:settings>'
          ignoring case.
      endif.

      find 'hideGrammaticalErrors' in l_xml.
      if sy-subrc ne 0.
        replace first occurrence of '</w:settings>'
          in l_xml
          with '<w:hideGrammaticalErrors/></w:settings>'
          ignoring case.
      endif.

      set_file(
        i_path = 'word/settings.xml'
        i_text = l_xml ).

    endif.

  endmethod.                    "set_properties


  method set_docxml.

    doc_xml = i_xml.

  endmethod.                    "set_DOC_XML


  method set_file.

    read table r_zip->files transporting no fields
      with key
        name = i_path.
    if sy-subrc eq 0.
      delete_file( i_path ).
    endif.

    r_zip->add(
      name    = i_path
      content = zcl_convert_static=>text2xtext( i_text ) ).

  endmethod.                    "set_file


  method set_header_footer.

    data : l_filename     type string,
           l_string       type string,
           l_file         type string,
           l_xml          type string,
           l_xmlns        type string,
           l_style        type string,
           ls_list_object like line of t_list_object.

    if type ne zcl_word_static=>c_type_header and type ne zcl_word_static=>c_type_footer.
      return.
    endif.

* Build header/footer xml fragment
    l_xml =
      zcl_word_static=>build_text(
        i_text   = textline
        i_style  = style
        is_style = style_effect ).

    clear l_style.
    if not line_style is initial or not line_style_effect is initial.

      l_style =
        zcl_word_static=>build_paragraph_style(
          i_style  = line_style
          is_style = line_style_effect
        i_section = section_xml ).

      clear section_xml.

    endif.

    if type = zcl_word_static=>c_type_header.
      l_string = 'hdr'.
    else.
      l_string = 'ftr'.
    endif.
    concatenate '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:'
                l_string
                ns_xml
                '>'
                '<w:p>'
                l_style
                l_xml
                '</w:p>'
                '</w:'
                l_string
                '>'
                into l_xml respecting blanks.

* Search available header/footer name
    do.
      l_filename = sy-index.
      concatenate 'word/'
                  type
                  l_filename
                  '.xml'
                  into l_filename.
      condense l_filename no-gaps.

      read table r_zip->files with key name = l_filename
                 transporting no fields.
      if sy-subrc ne 0.
        exit. "exit do
      endif.
    enddo.

* Add header/footer file into zip
    set_file(
      i_path = l_filename
      i_text = l_xml ).

* Add content type exception for new header/footer
    l_file = get_file( '[Content_Types].xml' ).

    concatenate '<Override PartName="/'
                l_filename
                '" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.'
                type
                '+xml"/>'
                '</Types>'
                into l_string.
    replace '</Types>' with l_string into l_file.

* Update content type file
    set_file(
      i_path = '[Content_Types].xml'
      i_text = l_file ).

* Get relation file
    l_file = get_file( 'word/_rels/document.xml.rels' ).

* Create header/footer ID
    do.
      id = 'rId' && sy-index.                               "#EC NOTEXT
      l_string = 'Id="' && id && '"'.                       "#EC NOTEXT
      find first occurrence of l_string in l_file ignoring case.
      if sy-subrc ne 0.
        exit. "exit do
      endif.
    enddo.

* Update object list
    ls_list_object-id = id.
    ls_list_object-type = type.
    ls_list_object-path = l_filename.
    append ls_list_object to t_list_object.

* Add relation
    l_filename = l_filename+5.
    concatenate '<Relationship Id="'
                id
                '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
                type
                '" Target="'
                l_filename
                '"/>'
                '</Relationships>'
                into l_string.
    replace '</Relationships>' with l_string into l_file.

* Update relation file
    set_file(
      i_path = 'word/_rels/document.xml.rels'
      i_text = l_file ).

    if usenow_default = zcl_word_static=>c_true and type = zcl_word_static=>c_type_header.
      call method set_header_footer_direct
        exporting
          i_header = id.
    elseif usenow_default = zcl_word_static=>c_true and type = zcl_word_static=>c_type_footer.
      call method set_header_footer_direct
        exporting
          i_footer = id.
    endif.

    if usenow_first = zcl_word_static=>c_true and type = zcl_word_static=>c_type_header.
      call method set_header_footer_direct
        exporting
          i_header_first = id.
    elseif usenow_first = zcl_word_static=>c_true and type = zcl_word_static=>c_type_footer.
      call method set_header_footer_direct
        exporting
          i_footer_first = id.
    endif.

  endmethod.                    "write_headerfooter


  method set_header_footer_direct.

    if i_header is supplied.
      s_section-header = i_header.
    endif.

    if i_footer is supplied.
      s_section-footer = i_footer.
    endif.

    if i_header_first is supplied.
      s_section-header_first = i_header_first.
    endif.

    if i_footer_first is supplied.
      s_section-footer_first = i_footer_first.
    endif.

  endmethod.                    "header_footer_direct_assign


  method set_page_properties.

    " Define orientation
    if i_orientation = zcl_word_static=>c_orient_landscape.
      s_section-landscape = zcl_word_static=>c_true.
    endif.

    " Define Border
    if i_border_left is supplied.
      s_section-border_left = i_border_left.
    endif.

    if i_border_top is supplied.
      s_section-border_top = i_border_top.
    endif.

    if i_border_right is supplied.
      s_section-border_right = i_border_right.
    endif.

    if i_border_bottom is supplied.
      s_section-border_bottom = i_border_bottom.
    endif.

  endmethod.                    "set_params


  method set_title.

    set_document_properties(
      i_title = i_text ).

  endmethod.                    "set_title
ENDCLASS.
