/* APCSA 题库 · 2025 版 CED 知识点树
 * 来源：College Board AP Computer Science A Course and Exam Description (Effective Fall 2025)
 * 4 个 Unit / 53 个 Topic
 * 该文件以 JS 全局变量形式提供，使站点在 file:// 与 GitHub Pages 下均可离线运行
 */
window.QB_TOPICS = {
  ced: "2025 AP CSA Course and Exam Description（2025 秋季生效）",
  units: [
    {
      id: "U1",
      code: "Unit 1",
      name: "Using Objects and Methods",
      nameCn: "使用对象与方法",
      weight: "15–25%",
      color: "#2563eb",
      desc: "Java 程序基础、表达式与运算、方法调用、Math 类、对象与 String",
      topics: [
        { id: "1.1",  name: "Algorithms, Programming, and Compilers", nameCn: "算法、编程与编译器", kw: ["算法", "编译器", "ALU", "程序执行"] },
        { id: "1.2",  name: "Variables and Data Types", nameCn: "变量与数据类型", kw: ["int", "double", "boolean", "声明", "初始化"] },
        { id: "1.3",  name: "Expressions", nameCn: "表达式（输出 / 字符串字面量 / 算术）", kw: ["System.out.print", "字符串", "算术运算", "运算符优先级", "整数除法"] },
        { id: "1.4",  name: "Assignment Statements and Input", nameCn: "赋值语句与输入", kw: ["赋值", "Scanner", "nextInt", "nextLine", "输入"] },
        { id: "1.5",  name: "Casting and Ranges of Variables", nameCn: "强制类型转换与数值范围", kw: ["cast", "强制转换", "溢出", "Integer.MAX_VALUE"] },
        { id: "1.6",  name: "Compound Assignment Operators", nameCn: "复合赋值运算符", kw: ["+=", "-=", "*=", "/=", "%=", "自增自减"] },
        { id: "1.7",  name: "API and Libraries", nameCn: "API 与库", kw: ["API", "库", "文档"] },
        { id: "1.8",  name: "Comments and Documentation", nameCn: "注释与文档", kw: ["注释", "//", "/* */", "文档"] },
        { id: "1.9",  name: "Method Signatures", nameCn: "方法签名", kw: ["方法签名", "返回值", "参数", "重载"] },
        { id: "1.10", name: "Calling Class Methods", nameCn: "调用类方法", kw: ["类方法", "static", "Math.abs"] },
        { id: "1.11", name: "Math Class", nameCn: "Math 类", kw: ["Math.random", "Math.abs", "Math.pow", "Math.sqrt"] },
        { id: "1.12", name: "Objects and Classes", nameCn: "对象与类", kw: ["对象", "类", "实例", "属性"] },
        { id: "1.13", name: "Object Creation and Storage", nameCn: "对象创建与存储", kw: ["new", "构造", "引用", "对象变量"] },
        { id: "1.14", name: "Calling Instance Methods", nameCn: "调用实例方法", kw: ["实例方法", "点运算符", "返回值"] },
        { id: "1.15", name: "String Class", nameCn: "String 类", kw: ["String", "substring", "indexOf", "length", "equals", "compareTo", "不可变"] }
      ]
    },
    {
      id: "U2",
      code: "Unit 2",
      name: "Selection and Iteration",
      nameCn: "选择与迭代",
      weight: "25–35%",
      color: "#0891b2",
      desc: "布尔表达式、if / 嵌套 if、while / for 循环、嵌套循环、运行时间分析",
      topics: [
        { id: "2.1",  name: "Algorithms with Selection and Repetition", nameCn: "选择与重复的算法", kw: ["算法", "控制结构", "选择", "重复"] },
        { id: "2.2",  name: "Boolean Expressions", nameCn: "布尔表达式", kw: ["==", "!=", "<", ">", "equals", "比较"] },
        { id: "2.3",  name: "if Statements", nameCn: "if 语句", kw: ["if", "else", "else if", "单分支", "多分支"] },
        { id: "2.4",  name: "Nested if Statements", nameCn: "嵌套 if 语句", kw: ["嵌套", "nested if", "分层判断"] },
        { id: "2.5",  name: "Compound Boolean Expressions", nameCn: "复合布尔表达式", kw: ["&&", "||", "!", "短路求值", "德摩根"] },
        { id: "2.6",  name: "Comparing Boolean Expressions", nameCn: "布尔表达式的比较与等价变换", kw: ["等价", "德摩根定律", "真值表"] },
        { id: "2.7",  name: "while Loops", nameCn: "while 循环", kw: ["while", "哨兵循环", "循环条件", "无限循环"] },
        { id: "2.8",  name: "for Loops", nameCn: "for 循环", kw: ["for", "初始化", "循环条件", "更新", "步长"] },
        { id: "2.9",  name: "Implementing Selection and Iteration Algorithms", nameCn: "选择与循环算法实现", kw: ["算法实现", "累加", "计数", "最值"] },
        { id: "2.10", name: "Implementing String Algorithms", nameCn: "字符串算法实现", kw: ["字符串遍历", "substring", "charAt", "拼接"] },
        { id: "2.11", name: "Nested Iteration", nameCn: "嵌套循环", kw: ["嵌套循环", "双重循环", "输出模式"] },
        { id: "2.12", name: "Informal Run-Time Analysis", nameCn: "非形式化运行时间分析", kw: ["复杂度", "O(n)", "循环次数", "算法效率"] }
      ]
    },
    {
      id: "U3",
      code: "Unit 3",
      name: "Class Creation",
      nameCn: "类的创建",
      weight: "10–18%",
      color: "#7c3aed",
      desc: "抽象与类设计、构造器、方法编写、引用传递、static、作用域、this",
      topics: [
        { id: "3.1", name: "Abstraction and Program Design", nameCn: "抽象与程序设计", kw: ["抽象", "封装", "类设计"] },
        { id: "3.2", name: "Impact of Program Design", nameCn: "程序设计的影响", kw: ["可维护性", "复用", "设计影响"] },
        { id: "3.3", name: "Anatomy of a Class", nameCn: "类的解剖", kw: ["属性", "实例变量", "private", "public", "方法"] },
        { id: "3.4", name: "Constructors", nameCn: "构造器", kw: ["构造器", "constructor", "重载", "默认构造器"] },
        { id: "3.5", name: "Methods: How to Write Them", nameCn: "方法的写法", kw: ["方法定义", "返回类型", "参数", "void"] },
        { id: "3.6", name: "Methods: Passing and Returning References of an Object", nameCn: "引用传递与返回", kw: ["传引用", "返回对象", "别名"] },
        { id: "3.7", name: "Class Variables and Methods", nameCn: "类变量与方法（static）", kw: ["static", "类变量", "类方法"] },
        { id: "3.8", name: "Scope and Access", nameCn: "作用域与访问", kw: ["作用域", "局部变量", "shadowing", "可见性"] },
        { id: "3.9", name: "this Keyword", nameCn: "this 关键字", kw: ["this", "实例引用", "消除歧义"] }
      ]
    },
    {
      id: "U4",
      code: "Unit 4",
      name: "Data Collections",
      nameCn: "数据集合",
      weight: "30–40%",
      color: "#dc2626",
      desc: "数组 / ArrayList / 二维数组、文件读入、包装类、搜索排序、递归",
      topics: [
        { id: "4.1",  name: "Data Collection Ethics", nameCn: "数据收集的伦理与社会影响", kw: ["伦理", "隐私", "数据安全"] },
        { id: "4.2",  name: "Introduction to Data Sets", nameCn: "数据集使用入门", kw: ["数据集", "数据分析"] },
        { id: "4.3",  name: "1D Array Creation and Access", nameCn: "一维数组的创建与访问", kw: ["数组", "length", "下标", "默认值", "ArrayIndexOutOfBounds"] },
        { id: "4.4",  name: "Array Traversal", nameCn: "数组遍历", kw: ["遍历", "增强 for", "for-each", "下标遍历"] },
        { id: "4.5",  name: "Array Algorithms", nameCn: "数组算法实现", kw: ["累加", "最值", "查找", "模式"] },
        { id: "4.6",  name: "Text File Reading", nameCn: "文本文件读入", kw: ["File", "Scanner", "hasNext", "nextLine", "文件读入"] },
        { id: "4.7",  name: "Wrapper Classes", nameCn: "包装类", kw: ["Integer", "Double", "parseInt", "parseDouble", "自动装箱"] },
        { id: "4.8",  name: "ArrayList Methods", nameCn: "ArrayList 方法", kw: ["ArrayList", "add", "remove", "set", "get", "size"] },
        { id: "4.9",  name: "ArrayList Traversal", nameCn: "ArrayList 遍历", kw: ["ArrayList 遍历", "遍历中删除", "下标错位"] },
        { id: "4.10", name: "ArrayList Algorithms", nameCn: "ArrayList 算法实现", kw: ["ArrayList 算法", "筛选", "插入", "remove"] },
        { id: "4.11", name: "2D Array Creation and Access", nameCn: "二维数组的创建与访问", kw: ["二维数组", "行", "列", "length", "初始化"] },
        { id: "4.12", name: "2D Array Traversal", nameCn: "二维数组遍历", kw: ["二维遍历", "行主序", "嵌套 for"] },
        { id: "4.13", name: "2D Array Algorithms", nameCn: "二维数组算法实现", kw: ["二维算法", "行列统计", "矩阵"] },
        { id: "4.14", name: "Searching Algorithms", nameCn: "搜索算法", kw: ["线性查找", "二分查找", "顺序查找"] },
        { id: "4.15", name: "Sorting Algorithms", nameCn: "排序算法", kw: ["选择排序", "插入排序", "归并排序", "排序"] },
        { id: "4.16", name: "Recursion Basics", nameCn: "递归基础", kw: ["递归", "基准情形", "递归调用", "栈"] },
        { id: "4.17", name: "Recursive Searching and Sorting", nameCn: "递归搜索与排序", kw: ["递归二分", "递归归并", "递归"] }
      ]
    }
  ]
};

/* 旧 CED（2019–2024）Unit(1–10) → 新 CED（2025）Unit(1–4) 映射
 * 用于把 2022–2024 年真题按旧大纲标注的题目归并到新大纲
 */
window.QB_OLD_UNIT_MAP = {
  "Unit 1":  { new: "U1", desc: "Primitive Types → 归入 Unit 1/2" },
  "Unit 2":  { new: "U1", desc: "Using Objects → Unit 1" },
  "Unit 3":  { new: "U2", desc: "Boolean Expressions and if Statements → Unit 2" },
  "Unit 4":  { new: "U2", desc: "Iteration → Unit 2" },
  "Unit 5":  { new: "U3", desc: "Writing Classes → Unit 3" },
  "Unit 6":  { new: "U4", desc: "Array → Unit 4" },
  "Unit 7":  { new: "U4", desc: "ArrayList → Unit 4" },
  "Unit 8":  { new: "U4", desc: "2D Array → Unit 4" },
  "Unit 9":  { new: "U4", desc: "Inheritance → 新版已并入 Unit 3/4" },
  "Unit 10": { new: "U4", desc: "Recursion → Unit 4" }
};
