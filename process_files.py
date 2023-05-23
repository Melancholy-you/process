import os
import re
import json
import mammoth
import argparse


def clean_text(text):
    # 保留所有字母数字字符，删除其他所有字符
    text = re.sub(r'\W+', '', text)
    return text


def process_files(input_dir, output_dir):  # 输入目录路径，输出目录路径
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)  # 如果目录不存在，创建目录

    # 遍历输入目录下的所有文件
    for filename in os.listdir(input_dir):  # os.listdir() 方法用于返回指定的文件夹包含的文件或文件夹的名字的列表
        # 我们只处理.docx文件
        if filename.endswith('.docx'):  # 如果文件名以.docx结尾
            input_path = os.path.join(input_dir, filename)  # os.path.join() 方法用于路径拼接文件路径
            output_path = os.path.join(output_dir, filename.replace('.docx', '.json'))  # 将.docx替换为.json

            # 从.docx文件读取文本
            with open(input_path, 'rb') as docx_file:  # 以二进制读取文件
                text = mammoth.extract_raw_text(docx_file).value  # 读取文本

            # 清理文本  保留所有字母数字字符，删除其他所有字符
            cleaned_text = clean_text(text)

            # 创建一个字典
            data_dict = {"text": cleaned_text}  # 字典的键值对

            # 将字典写入json文件
            with open(output_path, 'w', encoding='utf-8') as file:  # 以写入的方式打开文件
                json.dump(data_dict, file)  # 将字典写入文件


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Process some files.')  # 创建解析器
    parser.add_argument('input_dir', type=str, help='Input directory path')  # 添加参数
    parser.add_argument('output_dir', type=str, help='Output directory path')  # 添加参数
    args = parser.parse_args()  # 解析参数
    print("处理中...")
    # 处理文件
    process_files(args.input_dir, args.output_dir)  # 调用函数
    print("处理完成！")

# ================================================================= #
#  Rust调用脚本模板
# use
# std::process::Command;
#
# fn
# main()
# {
#     let
# output = Command::new("python3") \
#     .arg("/path/to/your/python_script.py") \  # 更改为process_files.py
#     .arg("/path/to/input_directory") \        # 输入目录路径
#     .arg("/path/to/output_directory") \       # 输出目录路径
#     .output() \
#     .expect("Failed to execute command");
#
# if !output.status.success() {
# println!("Python script returned error:\n{}", String::
#     from_utf8_lossy( & output.stderr));
# }
# }
