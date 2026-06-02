import sys
import json
from docxtpl import DocxTemplate

def render_word_template(template_path, json_path, output_path):
    print(f"启动高保真模板渲染...\n- 模板: {template_path}\n- 数据: {json_path}")
    
    try:
        # 1. 加载模板文档
        doc = DocxTemplate(template_path)
        
        # 2. 读取 JSON 数据字典
        with open(json_path, 'r', encoding='utf-8') as f:
            context = json.load(f)
            
        # 3. 施展魔法：用数据渲染模板，自动保留所有局部格式！
        doc.render(context)
        
        # 4. 保存最终结果
        doc.save(output_path)
        print(f"文件已生成至: {output_path}")

    except Exception as e:
        print(f"发生错误: {e}")

if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("用法: python render_template.py <带{{标签}}的模板.docx> <info.json> <输出文档.docx>")
        sys.exit(1)
        
    render_word_template(sys.argv[1], sys.argv[2], sys.argv[3])