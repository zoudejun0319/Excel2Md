"""
测试Excel转Markdown工具
"""

from excel2md import Excel2Markdown


def test_basic_conversion():
    """测试基本转换功能"""
    print("=" * 50)
    print("测试1: 基本转换（所有工作表）")
    print("=" * 50)

    converter = Excel2Markdown('example.xlsx')
    markdown = converter.convert()

    print("\n生成的Markdown内容:\n")
    print(markdown)
    print("\n" + "=" * 50)

    # 保存到文件
    converter.save(markdown)


def test_single_sheet():
    """测试单个工作表转换"""
    print("\n" + "=" * 50)
    print("测试2: 转换单个工作表")
    print("=" * 50)

    converter = Excel2Markdown(
        'example.xlsx',
        output_path='output_single.md',
        sheet_name='产品销售'
    )

    markdown = converter.convert()
    print("\n生成的Markdown内容:\n")
    print(markdown)
    print("\n" + "=" * 50)

    converter.save(markdown)


if __name__ == "__main__":
    try:
        # 首先创建示例文件
        print("正在创建示例Excel文件...\n")
        import create_example
        create_example.create_example_data()

        print("\n开始测试转换功能...\n")

        # 运行测试
        test_basic_conversion()
        test_single_sheet()

        print("\n✓ 所有测试完成!")
        print("  - example.md (所有工作表)")
        print("  - output_single.md (单个工作表)")

    except Exception as e:
        print(f"\n✗ 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
