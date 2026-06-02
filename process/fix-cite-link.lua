-- 为 citeproc 生成的文献引用添加超链接跳转到参考文献条目。
-- 必须在 --citeproc 之后运行（citeproc 渲染 [N] → 此过滤器加超链接）。
--
-- citeproc 后的 AST：
--   Cite{id="key"} [ Superscript [ Str "[N]" ] ]     ← 正文引用
--   Div[id="ref-key", class="csl-entry"] ...          ← 参考文献条目
--
-- 处理：
--   1. 为每个参考文献条目插入区间书签（bookmarkStart … bookmarkEnd）
--   2. 将正文 Cite 替换为指向该书签的超链接

local bkmk_counter = 7000

-- 为参考文献条目添加行内区间书签（移除 Div 的 id 避免生成块级书签）
local function add_entry_bookmark(el)
  if el.t ~= "Div" then return nil end
  if not el.identifier or el.identifier == "" then return nil end
  if not el.classes:includes("csl-entry") then return nil end

  local bkmk_name = el.identifier  -- citeproc 已生成 ref-citekey 格式的 id

  -- 清除 Div id，防止 pandoc 生成块级书签（块级书签点击不选中文本）
  el.identifier = ""

  bkmk_counter = bkmk_counter + 1
  local b_id = tostring(bkmk_counter)

  local bkmk_start = pandoc.RawInline("openxml",
    '<w:bookmarkStart w:id="' .. b_id .. '" w:name="' .. bkmk_name .. '"/>')
  local bkmk_end = pandoc.RawInline("openxml",
    '<w:bookmarkEnd w:id="' .. b_id .. '"/>')

  -- 在条目的第一个段落内容首尾插入行内书签，使 Word 点击后可选中整段
  if el.content and #el.content > 0 then
    local first_block = el.content[1]
    if first_block.t == "Para" or first_block.t == "Plain" then
      local new_content = pandoc.List()
      table.insert(new_content, bkmk_start)
      for _, inline in ipairs(first_block.content) do
        table.insert(new_content, inline)
      end
      table.insert(new_content, bkmk_end)
      first_block.content = new_content
    end
  end

  return el
end

-- 将正文 Cite 引用替换为超链接
local function fix_body_cite(el)
  if el.t ~= "Cite" then return nil end
  if not el.citations or #el.citations == 0 then return nil end

  local cite_id = el.citations[1].id
  if not cite_id or cite_id == "" then return nil end

  local bkmk_name = "ref-" .. cite_id
  local num_text = pandoc.utils.stringify(el):gsub("[%[%]]", "")

  return pandoc.RawInline("openxml",
    '<w:hyperlink w:anchor="' .. bkmk_name .. '" w:history="1">' ..
    '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>' ..
    '<w:t>[' .. num_text .. ']</w:t></w:r>' ..
    '</w:hyperlink>')
end

return {
  { Div = add_entry_bookmark },
  { Cite = fix_body_cite }
}
