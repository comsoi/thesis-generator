-- 将表题从 Figure/Table AST 节点中提取为独立段落，置于内容之前。
-- 中文学术规范：表题在上，图题在下。

local function is_table_figure(el)
  if el.identifier and (el.identifier:match("^tab:") or el.identifier:match("^tbl:")) then
    return true
  end
  local found = false
  pandoc.walk_block(el, { Table = function() found = true end })
  return found
end

local function extract_caption_para(el)
  local caption = el.caption
  if not caption or not caption.long or #caption.long == 0 then
    return nil
  end
  local blk = caption.long[1]
  local p
  if blk.t == "Plain" then
    p = pandoc.Para(blk.content)
  else
    p = pandoc.Para(blk.content)
  end
  -- pandoc 3.x 中 Para 不支持直接设 attr，通过 Div 包裹传递 custom-style 到 DOCX
  return pandoc.Div({p}, pandoc.Attr("", {}, {{"custom-style", "Table Caption"}}))
end

-- Figure handler
local function fix_figure(el)
  if not is_table_figure(el) then
    return nil
  end
  local caption_para = extract_caption_para(el)
  if not caption_para then
    return nil
  end
  el.caption = pandoc.Caption({})
  return { caption_para, el }
end

-- Table handler: 处理独立的 Table 节点（无 #figure 包装但有 caption 的情况）
local function fix_table(el)
  local caption_para = extract_caption_para(el)
  if not caption_para then
    return nil
  end
  el.caption = pandoc.Caption({})
  return { caption_para, el }
end

return {
  { Figure = fix_figure, Table = fix_table }
}
