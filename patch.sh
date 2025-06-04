#!/usr/bin/env docx-patcher.sh
# shellcheck disable=SC2120
# shellcheck shell=bash

cur_xpath=

:doc:enter_sdt() {
  local name="${1:?no name}" escaped_name
  escaped_name="$(xml_escape_string "$name")"
  :xml.push_xpath "//w:sdt[w:sdtPr/w:docPartObj/w:docPartGallery[@w:val=\"$escaped_name\"]]"
}

:doc:leave() {
  :xml.pop_xpath "$@";
}

:xml.file "word/document.xml" && {
  :xml.edit -d '/w:document/w:body/*[1][self::w:p][w:r/w:br]'

  :doc:enter_sdt 'Table of Contents'
    :xml.push_xpath "$cur_xpath/w:sdtContent/w:p[1]"
      :xml.edit -d "\$cur/w:r"
      __edit_args=()
      ,begin "$cur_xpath"
        ,subnode w:r
        ,begin
          ,subnode w:rPr
          ,begin
            ,subnode w:rFonts @w:hint=eastAsia
            ,subnode w:fitText @w:val=1083
            ,subnode w:spacing @w:val=360
          ,end
          ,subnode w:t 'text()'=目
        ,end

        ,subnode w:r
        ,begin
          ,subnode w:rPr
          ,begin
            ,subnode w:rFonts @w:hint=eastAsia
            ,subnode w:fitText @w:val=1083
            ,subnode w:spacing @w:val=1
          ,end
          ,subnode w:t 'text()'=录
        ,end
      ,end
      :xml.edit "${__edit_args[@]}"
      __edit_args=()
    :xml.pop_xpath

    :xml.push_xpath "$cur_xpath/w:sdtContent/w:p[2]"
      :xml.edit -d "\$cur/w:r"
      __edit_args=()
      ,begin "$cur_xpath"
        ,subnode w:r
        ,begin
          ,subnode w:fldChar @w:fldCharType=begin @w:dirty=true
          ,subnode w:instrText @xml:space=preserve 'text()'=\
'TOC \h \z \t "标题 1;1;标题 2;2;标题 3;3;References Heading;1;Acknowledgements Heading;1"'
          ,subnode w:fldChar @w:fldCharType=separate
          ,subnode w:fldChar @w:fldCharType=end
        ,end
      ,end
      :xml.edit "${__edit_args[@]}"
      __edit_args=()
    :xml.pop_xpath
  :doc:leave

  :doc:enter_sdt 'List of Figures'
    :xml.push_xpath "$cur_xpath/w:sdtContent/w:p[1]"
      :xml.edit -d "\$cur/w:r"
      __edit_args=()
      ,begin "$cur_xpath"
        ,subnode w:r
        ,begin
          ,subnode w:rPr
          ,begin
            ,subnode w:rFonts \
              @w:hint=eastAsia
          ,end
          ,subnode w:t \
            @xml:space=preserve \
            'text()'=插图清单
        ,end
      ,end
      :xml.edit "${__edit_args[@]}"
      __edit_args=()
    :xml.pop_xpath
  :doc:leave

  :doc:enter_sdt 'List of Tables'
    :xml.push_xpath "$cur_xpath/w:sdtContent/w:p[1]"
      :xml.edit -d "\$cur/w:r"
      __edit_args=()
      ,begin "$cur_xpath"
        ,subnode w:r
        ,begin
          ,subnode w:rPr
          ,begin
            ,subnode w:rFonts \
              @w:hint=eastAsia
          ,end
          ,subnode w:t \
            @xml:space=preserve \
            'text()'=表格清单
        ,end
      ,end
      :xml.edit "${__edit_args[@]}"
      __edit_args=()
    :xml.pop_xpath
  :doc:leave

  :xml.push_xpath '/w:document/w:body/w:p[w:pPr[w:pStyle[@w:val="ImageCaption"]]]'
    :xml.push_xpath "$cur_xpath/preceding-sibling::*[1][self::w:tbl]"
      :style.table:begin
        :style.table.borders:begin
          :style.table.borders.at top _type=none
          :style.table.borders.at bottom _type=none
        :style:end
      :style:end
      :xml.push_xpath "$cur_xpath/w:tr/w:tc/w:tbl"
        :style.table:begin
          :style.table.borders:begin
            :style.table.borders.at top _type=none
            :style.table.borders.at bottom _type=none
          :style:end
        :style:end
        :xml.edit -u "\$cur//w:pStyle[@w:val=\"ImageCaption\"]/@w:val" -v 'Compact'
      :xml.pop_xpath
    :xml.pop_xpath
  :xml.pop_xpath
}
