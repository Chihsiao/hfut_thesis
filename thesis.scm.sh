#!/usr/bin/env docx-patcher.sh
# shellcheck disable=SC1091
# shellcheck shell=bash

: "${EQ_COMPAT=1}"

source "$OOXML_PATCHER_ROOT/examples/normalize.scm.sh"
source "$OOXML_PATCHER_ROOT/examples/zh_CN.scm.sh"

:xml.file word/theme/theme1.xml && {
  :xml.edit --var 'fontScheme' '/a:theme/a:themeElements/a:fontScheme[@name="Office"]' \
    -u "\$fontScheme//a:font[@script='Hans']/@typeface" -v '宋体' \
    -u "\$fontScheme//a:latin/@typeface" -v 'Times New Roman'
}

:xml.file word/styles.xml && {
  :style:for 'Normal'
    :style.para:begin
      :style.para.spacing _before=0L _after=0L _line=1L
    :style:end
  :style:end

  :style:for 'Compact'
    :style.para:begin
      :style.para.justifying _to=both
      :style.para.spacing _before=0pt _after=0pt _line=22pt
    :style:end
  :style:end

  :style:for 'Body Text'
    :style.para:begin
      :style.para.spacing _before=0pt _after=0pt _line=22pt _lineRule=atLeast _contextual=true
    :style:end
  :style:end

  __edit_args=()
  ,begin '/w:styles'
    ,subnode w:style \
      @w:type=paragraph \
      @w:styleId=Equation \
      @w:customStyle=1
    ,begin
      ,subnode w:name @w:val='Equation'
      ,subnode w:basedOn @w:val=BodyText
      ,subnode w:next @w:val=Equation
      ,subnode w:qFormat
    ,end
  ,end
  :xml.edit "${__edit_args[@]}"
  __edit_args=()

  :style:for 'Equation'
    :style.para:begin
      :style.para.spacing _before=1.8pt _after=1.8pt _line=1L _contextual=true
    :style:end
  :style:end

#region heading
  :style:for 'Abstract Title'
    :style.para:begin
      :style.para.justifying _to=center
      :style.para.outline_level _level=0
      # if not to enable equivalent compatibility
      if ! bool "${EQ_COMPAT:-0}"
      then
        :style.para.spacing _before=0.5L _after=1.5L _line=1L
       else
        :style.para.spacing _before=0L _after=1.38L _line=31.5pt
      fi
    :style:end
    :style.run:begin
      :style.run.capital _caps=1
      :style.run.font_family _hAnsiFont='Times New Roman' _eastAsiaFont=宋体
      :style.run.font_size _size=小二
      :style.run.italic _italic=0
      :style.run.bold _bold=1
    :style:end
  :style:end

  :style:for 'TOC Heading'
    :style.para:begin
      :style.para.justifying _to=center
      :style.para.outline_level _level=0
      # if not to enable equivalent compatibility
      if ! bool "${EQ_COMPAT:-0}"
      then
        :style.para.spacing _before=0.5L _after=1.5L _line=1L
       else
        :style.para.spacing _before=0.5L _after=1.5L _line=22pt
      fi
    :style:end
    :style.run:begin
      :style.run.font_size _size=小二
      :style.run.font_family _hAnsiFont='Times New Roman' _eastAsiaFont=宋体
      :style.run.italic _italic=0
      :style.run.bold _bold=1
    :style:end
  :style:end

  :style:for 'Heading 1'
    :style.para:begin
      :style.para.justifying _to=center
      # if not to enable equivalent compatibility
      if ! bool "${EQ_COMPAT:-0}"
      then
        :style.para.spacing _before=1L _after=1L _line=22pt
       else
        :style.para.spacing _before=0L _after=0.76L _line=42pt
      fi
    :style:end
    :style.run:begin
      :style.run.font_family _hAnsiFont=黑体 _eastAsiaFont=黑体
      :style.run.font_size _size=三号
      :style.run.italic _italic=0
      :style.run.bold _bold=1
    :style:end
  :style:end

  :style:for 'Heading 2'
    :style.para:begin
      :style.para.justifying _to=start
      :style.para.spacing _before=0.5L _after=0.5L _line=22pt
    :style:end
    :style.run:begin
      :style.run.font_family _hAnsiFont=黑体 _eastAsiaFont=黑体
      :style.run.font_size _size=小四
      :style.run.italic _italic=0
      :style.run.bold _bold=0
    :style:end
  :style:end

  :style:for 'Heading 3'
    :style.para:begin
      :style.para.justifying _to=start
      :style.para.spacing _before=0.5L _after=0.5L _line=22pt
    :style:end
    :style.run:begin
      :style.run.font_family _hAnsiFont='Times New Roman' _eastAsiaFont=宋体
      :style.run.font_size _size=小四
      :style.run.italic _italic=0
      :style.run.bold _bold=1
    :style:end
  :style:end

  __edit_args=()
  ,begin '/w:styles'
    ,subnode w:style \
      @w:type=paragraph \
      @w:styleId=ReferencesHeading \
      @w:customStyle=1
    ,begin
      ,subnode w:name @w:val='References Heading'
      ,subnode w:basedOn @w:val=Heading1
      ,subnode w:qFormat
    ,end
  ,end
  :xml.edit "${__edit_args[@]}"
  __edit_args=()

  :style:for 'References Heading'
    :style.para:begin
      if ! bool "${EQ_COMPAT:-0}" # not to enable
        then :style.para.spacing _before=0.5L _after=0.5L _line=22pt
        else :style.para.spacing _before=0L _after=0.38L _line=32pt
      fi
    :style:end
  :style:end

  __edit_args=()
  ,begin '/w:styles'
    ,subnode w:style \
      @w:type=paragraph \
      @w:styleId=AcknowledgementsHeading \
      @w:customStyle=1
    ,begin
      ,subnode w:name @w:val='Acknowledgements Heading'
      ,subnode w:basedOn @w:val=Heading1
      ,subnode w:qFormat
    ,end
  ,end
  :xml.edit "${__edit_args[@]}"
  __edit_args=()

  :style:for 'Acknowledgements Heading'
    :style.para:begin
      if ! bool "${EQ_COMPAT:-0}" # not to enable
        then :style.para.spacing _before=0.5L _after=1.5L _line=1L
        else :style.para.spacing _before=0L _after=1.58L _line=37.75pt
      fi
    :style:end
    :style.run:begin
      :style.run.font_family _hAnsiFont='Times New Roman' _eastAsiaFont=宋体
      :style.run.font_size _size=小二
    :style:end
  :style:end
#endregion

#region custom paragraph
  __edit_args=()
  ,begin '/w:styles'
    ,subnode w:style \
      @w:type=paragraph \
      @w:styleId=Paragraph \
      @w:customStyle=1
    ,begin
      ,subnode w:name @w:val='Paragraph'
      ,subnode w:basedOn @w:val=BodyText
      ,subnode w:next @w:val=Paragraph
      ,subnode w:qFormat
    ,end
  ,end
  :xml.edit "${__edit_args[@]}"
  __edit_args=()

  :style:for 'Paragraph'
    :style.para:begin
      :style.para.justifying _to=both
      :style.para.indenting _firstLine=2C
      :style.para.spacing _before=0L _after=0L _line=22pt
    :style:end
    :style.run:begin
      :style.run.font_family _hAnsiFont='Times New Roman' _eastAsiaFont=宋体
      :style.run.font_size _size=小四
    :style:end
  :style:end

  __edit_args=()
  ,begin '/w:styles'
    ,subnode w:style \
      @w:type=paragraph \
      @w:styleId=Keywords \
      @w:customStyle=1
    ,begin
      ,subnode w:name @w:val='Keywords'
      ,subnode w:basedOn @w:val=BodyText
      ,subnode w:next @w:val=Keywords
      ,subnode w:qFormat
    ,end
  ,end
  :xml.edit "${__edit_args[@]}"
  __edit_args=()

  :style:for 'Keywords'
    :style.para:begin
      :style.para.justifying _to=start
      :style.para.indenting _firstLine=0C
      :style.para.spacing _before=1L _after=0L _line=22pt
    :style:end
    :style.run:begin
      :style.run.font_family _hAnsiFont='Times New Roman' _eastAsiaFont=宋体
      :style.run.font_size _size=小四
    :style:end
  :style:end

  __edit_args=()
  ,begin '/w:styles'
    ,subnode w:style \
      @w:type=paragraph \
      @w:styleId=LatinParagraph \
      @w:customStyle=1
    ,begin
      ,subnode w:name @w:val='Latin Paragraph'
      ,subnode w:next @w:val=LatinParagraph
      ,subnode w:basedOn @w:val=Paragraph
      ,subnode w:qFormat
    ,end
  ,end
  :xml.edit "${__edit_args[@]}"
  __edit_args=()

  :style:for 'Latin Paragraph'
    :style.para:begin
      :style.para.indenting _firstLine=1C
    :style:end
  :style:end

  __edit_args=()
  ,begin '/w:styles'
    ,subnode w:style \
      @w:type=paragraph \
      @w:styleId=FlushRight \
      @w:customStyle=1
    ,begin
      ,subnode w:name @w:val='Flush Right'
      ,subnode w:basedOn @w:val=BodyText
      ,subnode w:next @w:val=FlushRight
      ,subnode w:qFormat
    ,end
  ,end
  :xml.edit "${__edit_args[@]}"
  __edit_args=()

  :style:for 'Flush Right'
    :style.para:begin
      :style.para.justifying _to=right
      :style.para.spacing _before=0L _after=0L _line=22pt
    :style:end
    :style.run:begin
      :style.run.font_family _hAnsiFont='Times New Roman' _eastAsiaFont=宋体
      :style.run.font_size _size=小四
    :style:end
  :style:end
#endregion

#region caption, figure, table
  :style:for 'Caption'
    :style.para:begin
      :style.para.justifying _to=center
      :style.para.spacing _before=0.5L _after=0.5L _line=1L
    :style:end
    :style.run:begin
      :style.run.font_family _hAnsiFont='Times New Roman' _eastAsiaFont=宋体
      :style.run.font_size _size=五号
      :style.run.italic _italic=0
      :style.run.bold _bold=0
    :style:end
  :style:end

  :style:for 'Figure'
    :style.para:begin
      :style.para.justifying _to=center
      :style.para.spacing _before=0.5L _after=0L _line=1L
    :style:end
  :style:end

  :style:for 'Table'
    :style.table:begin
      :style.table.justifying _to=center
      :style.table.borders:begin
        :style.table.borders.at top _type=single _size=0.5pt
        :style.table.borders.at bottom _type=single _size=0.5pt
      :style:end
    :style:end

    :style@table:for firstRow
      :style.table:begin
        :style.table.justifying _unset=1
      :style:end
      :style.table_row:begin
        :style.table_row.justifying _unset=1
      :style:end
      :style.table_cell:begin
        :style.table_cell.aligning _to=top
      :style:end
    :style:end
  :style:end
#endregion

#region table of (contents, figures, tables)
  for ((i=1; i <= 3; ++i)) do
    :style:for_toc "$i"
      :style.para:begin
        :style.para.justifying _to=start
        :style.para.indenting _start=$((2*(i-1)))C
        :style.para.spacing _before=0L _after=0L _line=22pt
      :style:end
      :style.run:begin
        :style.run.font_family _hAnsiFont='Times New Roman' _eastAsiaFont=宋体
        :style.run.font_size _size=小四
        :style.run.italic _italic=0
        :style.run.bold _bold="$(
          [[ "$i" -eq 1 ]] \
            && echo 1 \
            || echo 0 \
        )"
      :style:end
    :style:end
  done

  :style:for_tof
    :style.para:begin
      :style.para.justifying _to=start
      :style.para.spacing _before=0L _after=0L _line=22pt
    :style:end
    :style.run:begin
      :style.run.font_family _hAnsiFont='Times New Roman' _eastAsiaFont=宋体
      :style.run.font_size _size=小四
      :style.run.italic _italic=0
      :style.run.bold _bold=0
    :style:end
  :style:end
#endregion

#region bibliography
  :style:for 'Bibliography'
    :style.para:begin
      :style.para.indenting _hanging=2.5C
      :style.para.spacing _before=0L _after=0L _line=20pt
      if ! bool "${EQ_COMPAT:-0}" # not to enable
        then :style.para.justifying _to=start
        else :style.para.justifying _to=both
      fi
    :style:end
    :style.run:begin
      :style.run.font_family _hAnsiFont='Times New Roman' _eastAsiaFont=宋体
      :style.run.font_size _size=五号
    :style:end
  :style:end
#endregion
}

cat - > word/header_hfut_thesis.xml << EOF
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:jc w:val="center" />
      <w:snapToGrid w:val="0" />
      <w:spacing w:after="0" w:line="360" w:lineRule="auto" />
      <w:pBdr>
        <w:bottom w:val="single" w:color="auto" w:sz="6" w:space="1" />
      </w:pBdr>
    </w:pPr>
    <w:r>
      <w:rPr>
        <w:sz w:val="21" />
        <w:szCs w:val="21" />
        <w:rFonts w:hint="eastAsia" />
      </w:rPr>
      <w:t>合肥工业大学毕业设计（论文）</w:t>
    </w:r>
  </w:p>
</w:hdr>
EOF

cat - > word/footer_hfut_thesis.xml << EOF
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:jc w:val="center" />
      <w:snapToGrid w:val="0" />
      <w:spacing w:after="0" w:line="360" w:lineRule="auto" />
    </w:pPr>
    <w:r>
      <w:fldChar w:fldCharType="begin" />
    </w:r>
    <w:r>
      <w:rPr>
        <w:sz w:val="21" />
        <w:szCs w:val="21" />
      </w:rPr>
      <w:instrText xml:space="preserve">PAGE   \* MERGEFORMAT</w:instrText>
    </w:r>
    <w:r>
      <w:fldChar w:fldCharType="end" />
    </w:r>
  </w:p>
</w:ftr>
EOF

:xml.file 'word/_rels/document.xml.rels' && {
  __edit_args=()
  ,begin '/r:Relationships'
    ,subnode Relationship @Id=header_hfut_thesis @Target=header_hfut_thesis.xml \
      @Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/header'
    ,subnode Relationship @Id=footer_hfut_thesis @Target=footer_hfut_thesis.xml \
      @Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer'
  ,end
  :xml.edit "${__edit_args[@]}"
  __edit_args=()
}

:xml.file '[Content_Types].xml' && {
  __edit_args=()
  ,begin '/t:Types'
    ,subnode Override @PartName='/word/header_hfut_thesis.xml' \
      @ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
    ,subnode Override @PartName='/word/footer_hfut_thesis.xml' \
      @ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'
  ,end
  :xml.edit "${__edit_args[@]}"
  __edit_args=()
}

:xml.file 'word/document.xml' && {
  :xml.proc '
    <xsl:template match="@*|node()">
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
        </xsl:copy>
    </xsl:template>

    <xsl:template match="/w:document/w:body">
      <xsl:copy>
        <xsl:copy-of select="@*" />
        <xsl:copy-of select="node()[not(self::w:sectPr)]" />
        <w:sectPr>
          <w:pgSz w:w="11906" w:h="16838" />
          <w:pgMar
            w:top="1701"
            w:left="1588"
            w:right="1588"
            w:bottom="1701"
            w:header="851"
            w:footer="851"
            w:gutter="0"
          />
          <w:pgNumType w:start="1" />
          <w:docGrid w:type="lines" w:linePitch="312" w:charSpace="0" />
        </w:sectPr>
      </xsl:copy>
    </xsl:template>
  '

  __edit_args=()
  ,begin "/w:document/w:body/w:sectPr"
    ,subnode w:headerReference @w:type=default @r:id=header_hfut_thesis
    ,subnode w:footerReference @w:type=default @r:id=footer_hfut_thesis
  ,end
  :xml.edit "${__edit_args[@]}"
  __edit_args=()
}
