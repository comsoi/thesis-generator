-- ============================================================================
-- fix-ref.lua — Typst → DOCX 交叉引用与按章编号
--
-- 职责：
--   1. 交叉引用：将 @tab:/@fig:/@sec:/@eq: 转为可点击的 Word REF 域
--   2. 按章编号：图 3-1、表 3-2、公式(2-1)。章号由 Lua 计数器跟踪，
--      SEQ 域由 pass_reset_seq 注入的隐藏 { SEQ X \r 0 \h } 按章重置
--   3. 公式排版：用 3 列无边框表格实现公式居中 + 编号右对齐
--   4. 文献引用透传：非 tab/fig/sec/eq 前缀的 Cite 直接放行给 citeproc
-- ============================================================================

local bkmk_counter = 1000

local function clean_id(prefix, raw_id)
  local clean = raw_id:gsub("[^%w]", "")
  return string.sub(prefix .. clean, 1, 40)
end

local function fld_seq(seq_name)
  return pandoc.RawInline('openxml', '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ ' .. seq_name .. ' \\* ARABIC </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>')
end

-- 章号计数器（pass2 的 Header handler 在遇到 Ch1-5 标题时递增）
local chapter_num = 0

-- 生成"章号-{ SEQ }"格式的题注编号
-- 例如 chapter_num=3 时返回 { "3-", { SEQ Table \* ARABIC } }
-- 章号是静态文本，序号由 Word SEQ 域自动递增（按章重置由 pass_reset_seq 保证）
local function fld_chapter_seq(seq_name)
  return {
    pandoc.RawInline('openxml', '<w:r><w:t>' .. chapter_num .. '-</w:t></w:r>'),
    pandoc.RawInline('openxml', '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ ' .. seq_name .. ' \\* ARABIC </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'),
  }
end

local function fld_ref(bkmk_name, display_text, switches)
  switches = switches or "\\h"
  return pandoc.RawInline('openxml', '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF ' .. bkmk_name .. ' ' .. switches .. ' </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>' .. display_text .. '</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>')
end

local sec_levels = {}

local function sec_ref_prefix(sec_id)
  local level = sec_levels[sec_id] or sec_levels[sec_id:lower()]
  return level == nil or level > 1
end

local function append_tail_text(inlines, tail_text)
  if not tail_text or tail_text == "" then
    return
  end

  local leading_ws, rest = tail_text:match("^(%s+)(.*)$")
  if leading_ws then
    table.insert(inlines, pandoc.Space())
    if rest ~= "" then
      table.insert(inlines, pandoc.Str(rest))
    end
  else
    table.insert(inlines, pandoc.Str(tail_text))
  end
end

-- =========================================================
-- 正文引用拦截器
-- =========================================================
local function fix_refs(inlines)
  local new_inlines = {}
  local i = 1
  while i <= #inlines do
    local matched = false

    if inlines[i].t == "Cite" and inlines[i].citations and #inlines[i].citations > 0 then
      local cite_id = inlines[i].citations[1].id
      local full_ref_id = cite_id
      local punct = ""
      local consumes_next = false

      if inlines[i+1] and inlines[i+1].t == "Str" and inlines[i+1].text:sub(1,1) == ":" then
        local id_tail = inlines[i+1].text:sub(2):match("^([%w%-_]+)")
        if id_tail then
          full_ref_id = cite_id .. ":" .. id_tail
          punct = inlines[i+1].text:sub(2 + #id_tail)
          consumes_next = true
        end
      end

      local lower_id = full_ref_id:lower()
      if lower_id:match("^tab:") or lower_id:match("^tbl:") or lower_id:match("^fig:") or lower_id:match("^sec:") or lower_id:match("^eq:") then
        local prefix = lower_id:match("^tab:") and "T" or (lower_id:match("^fig:") and "F" or (lower_id:match("^sec:") and "S" or "E"))
        local label = lower_id:match("^tab:") and "Table" or (lower_id:match("^fig:") and "Figure" or "Section")
        local bkmk_name = clean_id(prefix, full_ref_id)

        io.stderr:write(string.format("[DEBUG 引用生成] 捕获源码 '%s' -> 寻找书签: %s\n", full_ref_id, bkmk_name))

        if #new_inlines > 0 then
          local prev = new_inlines[#new_inlines]
          if prev.t == "Str" and not prev.text:match("%s$") then table.insert(new_inlines, pandoc.Space())
          elseif prev.t ~= "Space" and prev.t ~= "LineBreak" and prev.t ~= "Str" then table.insert(new_inlines, pandoc.Space()) end
        end

        if lower_id:match("^eq:") then
          table.insert(new_inlines, pandoc.Str("Eq. ("))
          table.insert(new_inlines, fld_ref(bkmk_name, "0", "\\h"))
          table.insert(new_inlines, pandoc.Str(")"))
        elseif lower_id:match("^sec:") then
          if sec_ref_prefix(full_ref_id) then
            table.insert(new_inlines, pandoc.Str("Sec."))
            table.insert(new_inlines, pandoc.Space())
          end
          -- 章节引用使用段落编号，避免带出尾随分隔符
          table.insert(new_inlines, fld_ref(bkmk_name, "0", "\\n \\h"))
        else
          table.insert(new_inlines, fld_ref(bkmk_name, label, "\\h"))
        end

        append_tail_text(new_inlines, punct)
        if consumes_next then i = i + 1 end
        matched = true
      end
    end

    if not matched and inlines[i].t == "Str" then
      local full_str = inlines[i].text
      local before, raw_prefix, ref_tail, punct = full_str:match("(.*)@(%w+):([%w%-_]+)(.*)")
      if not raw_prefix then before, raw_prefix, ref_tail, punct = full_str:match("(.*)%[(%w+)%]:([%w%-_]+)(.*)") end

      if raw_prefix then
        local full_ref_id = raw_prefix .. ":" .. ref_tail
        local lower_id = full_ref_id:lower()
        if lower_id:match("^tab:") or lower_id:match("^tbl:") or lower_id:match("^fig:") or lower_id:match("^sec:") or lower_id:match("^eq:") then
          local prefix = lower_id:match("^eq:") and "E" or (lower_id:match("^sec:") and "S" or (lower_id:match("^fig:") and "F" or "T"))
          local label = lower_id:match("^eq:") and "Eq. (" or (lower_id:match("^sec:") and "Section" or (lower_id:match("^fig:") and "Figure" or "Table"))
          local bkmk_name = clean_id(prefix, full_ref_id)

          io.stderr:write(string.format("[DEBUG 引用生成] 捕获源码 '%s' -> 寻找书签: %s\n", full_ref_id, bkmk_name))

          if before and before ~= "" then
            table.insert(new_inlines, pandoc.Str(before))
            if not before:match("%s$") then
              table.insert(new_inlines, pandoc.Space())
            end
          elseif #new_inlines > 0 then
            table.insert(new_inlines, pandoc.Space())
          end

          if lower_id:match("^eq:") then
            table.insert(new_inlines, pandoc.Str("Eq. ("))
            table.insert(new_inlines, fld_ref(bkmk_name, "0", "\\h"))
            table.insert(new_inlines, pandoc.Str(")"))
          elseif lower_id:match("^sec:") then
            if sec_ref_prefix(full_ref_id) then
              table.insert(new_inlines, pandoc.Str("Sec."))
              table.insert(new_inlines, pandoc.Space())
            end
            table.insert(new_inlines, fld_ref(bkmk_name, "0", "\\n \\h"))
          else
            table.insert(new_inlines, fld_ref(bkmk_name, label, "\\h"))
          end
          append_tail_text(new_inlines, punct)
          matched = true
        end
      end
    end

    if not matched then table.insert(new_inlines, inlines[i]) end
    i = i + 1
  end
  return new_inlines
end

-- =========================================================
-- 源码级预处理：修正贪婪匹配，防止跨行吞噬
-- =========================================================
local source_content = ""
if PANDOC_STATE and PANDOC_STATE.input_files then
  for _, filename in ipairs(PANDOC_STATE.input_files) do
    local f = io.open(filename, "r")
    if f then source_content = source_content .. f:read("*all"); f:close() end
  end
end
local search_source = source_content:gsub("\\", ""):gsub("%s+", "")

local typst_table_labels = {}
local sec_fingerprints = {}

for m in source_content:gmatch("<(tab:[%w%-_:]+)>") do table.insert(typst_table_labels, m) end

-- 【修复点】：[^\n]- 确保只匹配当前行，阻止跨行读取垃圾内容
for heading_text, sec_id in source_content:gmatch("([^\n]-)<(sec:[%w%-_:]+)>") do
  local clean_before = heading_text:gsub("%s+", ""):gsub("[^%w]", ""):lower()
  if #clean_before > 0 then sec_fingerprints[string.sub(clean_before, -30)] = sec_id end
end
local source_search_pos = 1

-- =========================================================
-- 核心逻辑处理
-- =========================================================
local function process_math_block(el)
  local math_node, eq_label = nil, nil
  local text_length = 0

  pandoc.walk_block(el, {
    Math = function(m) if m.mathtype == "DisplayMath" then math_node = m elseif not math_node then math_node = m end end,
    Str = function(s) text_length = text_length + #s.text end
  })

  -- 修复：取第一个 eq: 标签，避免同一段落多标签时被最后一个覆盖
  pandoc.walk_block(el, {
    RawInline = function(raw) if not eq_label then local m = raw.text:match("<(eq:[%w%-_:]+)>") or raw.text:match("{#(eq:[%w%-_:]+)}"); if m then eq_label = m end end end,
    Str = function(str) if not eq_label then local m = str.text:match("<(eq:[%w%-_:]+)>") or str.text:match("{#(eq:[%w%-_:]+)}"); if m then eq_label = m end end end,
    Span = function(span) if not eq_label and span.identifier and span.identifier:match("^eq:") then eq_label = span.identifier end end,
    Code = function(c) if not eq_label then local m = c.text:match("<(eq:[%w%-_:]+)>") or c.text:match("{#(eq:[%w%-_:]+)}"); if m then eq_label = m end end end
  })
  if not eq_label and el.identifier and el.identifier:match("^eq:") then eq_label = el.identifier end

  -- 修复：源码搜索仅对 DisplayMath 启用、仅向前搜索，避免行内公式误匹配到前面公式的标签
  if math_node and not eq_label and search_source ~= "" and math_node.mathtype == "DisplayMath" then
    local plain_math = math_node.text:gsub("\\", ""):gsub("%s+", "")
    local math_suffix = string.sub(plain_math, -15)
    local start_pos, end_pos = string.find(search_source, math_suffix, source_search_pos, true)
    if end_pos then
      source_search_pos = end_pos
      local m = string.sub(search_source, end_pos + 1, end_pos + 40):match("<(eq:[%w%-_:]+)>") or string.sub(search_source, end_pos + 1, end_pos + 40):match("{#(eq:[%w%-_:]+)}")
      if m then eq_label = m end
    end
  end

  if math_node and (eq_label or (math_node.mathtype == "DisplayMath" and text_length < 20)) then
    math_node.mathtype = "DisplayMath"
    bkmk_counter = bkmk_counter + 1
    local b_id = tostring(bkmk_counter)
    local seq_fields = fld_chapter_seq("Equation")  -- 章节编号列表
    local bkmk_name = eq_label and clean_id("E", eq_label) or ("EqAuto" .. b_id)
    local bkmk_start = pandoc.RawInline('openxml', '<w:bookmarkStart w:id="' .. b_id .. '" w:name="' .. bkmk_name .. '"/>')
    local bkmk_end = pandoc.RawInline('openxml', '<w:bookmarkEnd w:id="' .. b_id .. '"/>')

    if eq_label then
      io.stderr:write(string.format("[DEBUG 公式实体生成] 成功抓取！绑定源标签: '%s' -> 实体书签: %s\n", eq_label, bkmk_name))
    end

    local md_table = "| | | |\n|:---|:---:|---:|\n| | EQUATION_X | NUMBER_Y |"
    local temp_doc = pandoc.read(md_table, "markdown")
    local tbl = temp_doc.blocks[1]

    if tbl.colspecs and #tbl.colspecs == 3 then
        tbl.colspecs[1] = { pandoc.AlignLeft, 0.10 }
        tbl.colspecs[2] = { pandoc.AlignCenter, 0.80 }
        tbl.colspecs[3] = { pandoc.AlignRight, 0.10 }
    end

    tbl = pandoc.walk_block(tbl, {
      Plain = function(blk)
        if #blk.content == 1 and blk.content[1].t == "Str" then
          if blk.content[1].text == "EQUATION_X" then
            return pandoc.Div({ pandoc.Para({ math_node }) }, pandoc.Attr("", {}, {{"custom-style", "FormulaPara"}}))
          elseif blk.content[1].text == "NUMBER_Y" then
            local eq_nums = pandoc.List({pandoc.Str("("), bkmk_start})
for _, s in ipairs(seq_fields) do eq_nums:insert(s) end
eq_nums:insert(bkmk_end)
eq_nums:insert(pandoc.Str(")"))
return pandoc.Div({ pandoc.Para(eq_nums) }, pandoc.Attr("", {}, {{"custom-style", "FormulaPara"}}))
          end
        end
      end,
      Para = function(blk)
        if #blk.content == 1 and blk.content[1].t == "Str" then
          if blk.content[1].text == "EQUATION_X" then
            return pandoc.Div({ pandoc.Para({ math_node }) }, pandoc.Attr("", {}, {{"custom-style", "FormulaPara"}}))
          elseif blk.content[1].text == "NUMBER_Y" then
            return pandoc.Div({ pandoc.Para(eq_nums) }, pandoc.Attr("", {}, {{"custom-style", "FormulaPara"}}))
          end
        end
      end
    })

    tbl.attr = tbl.attr or pandoc.Attr()
    tbl.attr.attributes["custom-style"] = "EquationTable"
    return tbl
  end
  return nil
end

local table_idx = 0
local pass1 = {
  Table = function(el)
    if el.caption and el.caption.long and #el.caption.long > 0 then
      table_idx = table_idx + 1
      if typst_table_labels[table_idx] then el.identifier = typst_table_labels[table_idx] end
    end
    return el
  end
}

-- pass_reset_seq：将隐藏 SEQ 重置域附加到章标题段落末尾，按章重置计数器。
--
-- 嵌在标题文字尾部，不产生额外段落。域格式：
--   { SEQ Table \r 0 \h } { SEQ Figure \r 0 \h } { SEQ Equation \r 0 \h }
--   \r 0  重置到 0（下一处 SEQ 首次显示 1）；\h  隐藏结果（F9 后不可见）
--
-- 跳过 Ch1（无前序图表无需重置），跳过 致谢/附录/References。
local pass_reset_seq = {
  Pandoc = function(doc)
    local seen_first = false
    for _, block in ipairs(doc.blocks) do
      if block.t == "Header" and block.level == 1 then
        local htext = pandoc.utils.stringify(block):gsub("%s+", "")
        if not htext:match("^致谢") and not htext:match("^附录")
           and not htext:match("^Acknowledg") and not htext:match("^Appendix")
           and not htext:match("^References") and not htext:match("^Bibliography") then
          if not seen_first then
            seen_first = true  -- Ch1，跳过
          else
            block.content:insert(pandoc.RawInline("openxml", '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ Table \\r 0 \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>0</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'))
            block.content:insert(pandoc.RawInline("openxml", '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ Figure \\r 0 \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>0</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'))
            block.content:insert(pandoc.RawInline("openxml", '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ Equation \\r 0 \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>0</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'))
          end
        end
      end
    end
    return doc
  end
}

local pass2 = {
  Header = function(el)
    -- 章节计数器：Heading 1 时递增 chapter_num，重置章内计数器（跳过后缀非正文标题）
    if el.level == 1 then
      local htext = pandoc.utils.stringify(el):gsub("%s+", "")
      if not htext:match("^致谢") and not htext:match("^附录") and not htext:match("^Acknowledg") and not htext:match("^Appendix") and not htext:match("^References") and not htext:match("^Bibliography") then
        chapter_num = chapter_num + 1
      end
    end

    local found_sec = nil

    -- 【修复点】：双保险策略。首选从 AST 自带的 identifier 中提取 (Pandoc 会自动识别 <sec:...>)
    if el.identifier and el.identifier:match("^sec:") then
      found_sec = el.identifier
    else
      -- 备用方案：降维纯文本指纹匹配
      local clean_text = pandoc.utils.stringify(el):gsub("%s+", ""):gsub("[^%w]", ""):lower()
      found_sec = sec_fingerprints[string.sub(clean_text, -30)]
    end

    if found_sec then
      while #el.content > 0 do
        local last = el.content[#el.content]
        if last.t == "Space" or last.t == "SoftBreak" or last.t == "LineBreak" then
          table.remove(el.content)
        else
          break
        end
      end

      sec_levels[found_sec] = el.level
      sec_levels[found_sec:lower()] = el.level
      bkmk_counter = bkmk_counter + 1
      local b_id = tostring(bkmk_counter)
      local bkmk_name = clean_id("S", found_sec)
      local bkmk_start = pandoc.RawInline('openxml', '<w:bookmarkStart w:id="' .. b_id .. '" w:name="' .. bkmk_name .. '"/>')
      local bkmk_end = pandoc.RawInline('openxml', '<w:bookmarkEnd w:id="' .. b_id .. '"/>')

      table.insert(el.content, 1, bkmk_start)
      table.insert(el.content, bkmk_end)

      -- 【新增点】：打印成功绑定书签的日志，方便你确认
      io.stderr:write(string.format("[DEBUG 章节实体生成] 成功绑定源标签: '%s' -> 实体书签: %s\n", found_sec, bkmk_name))
    end

    el.content = fix_refs(el.content)
    return el
  end,

  Para = function(el) local tbl = process_math_block(el); if tbl then return tbl end; el.content = fix_refs(el.content); return el end,
  Plain = function(el) local tbl = process_math_block(el); if tbl then return tbl end; el.content = fix_refs(el.content); return el end,

  Figure = function(el)
    local raw_id = el.identifier or ""
    local is_tab = raw_id:match("^tab:") or raw_id:match("^tbl:")
    if not is_tab and not raw_id:match("^fig:") then pandoc.walk_block(el, { Table = function() is_tab = true end }) end
    if raw_id == "" then return el end

    local prefix, label_str = is_tab and "T" or "F", is_tab and "Table" or "Figure"

    if el.caption and el.caption.long and #el.caption.long > 0 then
      local first_para = el.caption.long[1]
      if first_para.t == "Plain" then first_para = pandoc.Para(first_para.content) end

      bkmk_counter = bkmk_counter + 1
      local b_id = tostring(bkmk_counter)
      local bkmk_name = clean_id(prefix, raw_id)
      local bkmk_start = pandoc.RawInline('openxml', '<w:bookmarkStart w:id="' .. b_id .. '" w:name="' .. bkmk_name .. '"/>')
      local bkmk_end = pandoc.RawInline('openxml', '<w:bookmarkEnd w:id="' .. b_id .. '"/>')

      io.stderr:write(string.format("[DEBUG 图表实体生成] 成功绑定源标签: '%s' -> 实体书签: %s\n", raw_id, bkmk_name))

      table.insert(first_para.content, 1, pandoc.Str(": "))
      table.insert(first_para.content, 1, bkmk_end)
      -- 章节编号：STYLEREF-{SEQ}
      local seq_inlines = fld_chapter_seq(label_str)
      for j = #seq_inlines, 1, -1 do
        table.insert(first_para.content, 1, seq_inlines[j])
      end
      table.insert(first_para.content, 1, pandoc.Str(label_str .. " "))
      table.insert(first_para.content, 1, bkmk_start)

      el.caption.long[1] = first_para
    end
    if not is_tab and el.content and #el.content > 0 then
      el.content = { pandoc.Div(el.content, pandoc.Attr("", {}, {{"custom-style", "Figure"}})) }
    end
    return el
  end,

  Table = function(el)
    -- 修复：EquationTable 是公式编号表，不要覆盖其样式
    if el.attr and el.attr.attributes and el.attr.attributes["custom-style"] == "EquationTable" then
      return el
    end

    local raw_id = el.identifier or ""
    if raw_id == "" then bkmk_counter = bkmk_counter + 1; raw_id = "UnknownTab" .. bkmk_counter end

    if el.caption and el.caption.long and #el.caption.long > 0 then
      local first_para = el.caption.long[1]
      if first_para.t == "Plain" then first_para = pandoc.Para(first_para.content) end

      local new_inlines = pandoc.List()
      for _, inline in ipairs(first_para.content) do
        if inline.t == "Str" then
           local clean_text = inline.text:gsub("<[%w%-_:]+>", ""):gsub("{#[%w%-_:]+}", "")
           if clean_text ~= "" then new_inlines:insert(pandoc.Str(clean_text)) end
        else new_inlines:insert(inline) end
      end
      first_para.content = new_inlines

      bkmk_counter = bkmk_counter + 1
      local b_id = tostring(bkmk_counter)
      local bkmk_name = clean_id("T", raw_id)
      local bkmk_start = pandoc.RawInline('openxml', '<w:bookmarkStart w:id="' .. b_id .. '" w:name="' .. bkmk_name .. '"/>')
      local bkmk_end = pandoc.RawInline('openxml', '<w:bookmarkEnd w:id="' .. b_id .. '"/>')

      io.stderr:write(string.format("[DEBUG 表格实体生成] 成功绑定源标签: '%s' -> 实体书签: %s\n", raw_id, bkmk_name))

      table.insert(first_para.content, 1, pandoc.Str(": "))
      table.insert(first_para.content, 1, bkmk_end)
      local seq_inlines = fld_chapter_seq("Table")
      for j = #seq_inlines, 1, -1 do
        table.insert(first_para.content, 1, seq_inlines[j])
      end
      table.insert(first_para.content, 1, pandoc.Str("Table "))
      table.insert(first_para.content, 1, bkmk_start)

      el.caption.long[1] = first_para
    end

    if el.colspecs then for i = 1, #el.colspecs do el.colspecs[i][1] = pandoc.AlignCenter end end

    if not el.attr then el.attr = pandoc.Attr(raw_id) end
    el.attr.attributes["custom-style"] = "CenterTable"

    return el
  end,

}

-- pass0_cleanup_empty_paras：删除 pandoc Typst reader 在
-- 多级标题之间错误生成的空段落。
--
-- 根因：pandoc 3.9 的 Typst reader 将 <sec:xxx> 标签解析为独立
-- Para 块而非标题属性。当两个标题连续出现时，标签段落就被夹在
-- 中间，渲染为不可见的空行。
local pass0_cleanup_empty_paras = {
  Pandoc = function(doc)
    local new_blocks = {}

    for i, block in ipairs(doc.blocks) do
      if block.t == "Para" then
        local plain = pandoc.utils.stringify(block):gsub("%s+", "")
        if plain == "" then
          -- 空段落：检查前后是否临近 Header 块（跳过中间的空段落）
          local near_header = false

          -- 向前查找
          for j = i - 1, 1, -1 do
            if doc.blocks[j].t == "Header" then
              near_header = true
              break
            end
            local s = pandoc.utils.stringify(doc.blocks[j]):gsub("%s+", "")
            if s ~= "" then break end
          end

          -- 向后查找
          if not near_header then
            for j = i + 1, #doc.blocks do
              if doc.blocks[j].t == "Header" then
                near_header = true
                break
              end
              local s = pandoc.utils.stringify(doc.blocks[j]):gsub("%s+", "")
              if s ~= "" then break end
            end
          end

          if near_header then
            goto continue  -- 跳过此空段落
          end
        end
      end
      table.insert(new_blocks, block)
      ::continue::
    end
    doc.blocks = new_blocks
    return doc
  end
}

local pass3_merge_lists = {
  Pandoc = function(doc)
    local new_blocks = {}
    local prev_list = nil

    for _, block in ipairs(doc.blocks) do
      if block.t == "OrderedList" then
        if prev_list then
          -- 发现它前面紧挨着也是一个有序列表，直接把当前项塞进前一个列表里！
          for _, item in ipairs(block.content) do
            table.insert(prev_list.content, item)
          end
        else
          prev_list = block
          table.insert(new_blocks, block)
        end
      else
        -- 遇到非列表元素（如正文、标题），打断粘合
        prev_list = nil
        table.insert(new_blocks, block)
      end
    end
    doc.blocks = new_blocks
    return doc
  end
}



return { pass0_cleanup_empty_paras, pass1, pass_reset_seq, pass2, pass3_merge_lists }
