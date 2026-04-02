// 江汉大学篮球运动膝关节损伤问卷 - 模拟数据生成器
const fs = require('fs');

// 问卷结构定义
const questions = [
  { id: 1, title: "您的性别", type: "radio", options: ["男", "女"] },
  { id: 2, title: "您的身份", type: "radio", options: ["体育学院篮球专项", "校篮球队成员", "普通篮球爱好者（公体课/社团）"] },
  { id: 3, title: "您的篮球运动年限", type: "radio", options: ["1年以下", "1-3年", "3-5年", "5年以上"] },
  { id: 4, title: "您每周打篮球频率", type: "radio", options: ["1-2次", "3-4次", "5次及以上"] },
  { id: 5, title: "您在打篮球时是否受过膝关节损伤", type: "radio", options: ["是", "否"] },
  { id: 6, title: "您受过几次膝关节损伤", type: "radio", options: ["1次", "2次", "3次及以上"], condition: "Q5=是" },
  { id: 7, title: "您的损伤类型是", type: "checkbox", options: ["半月板损伤", "髌骨劳损（膝盖骨疼痛）", "韧带损伤（内侧/前交叉）", "髌腱炎（跳跃膝）", "其他/不清楚"], condition: "Q5=是" },
  { id: 8, title: "损伤发生在什么情况下", type: "radio", options: ["起跳落地", "急停变向", "防守滑步", "身体对抗", "其他"], condition: "Q5=是" },
  { id: 9, title: "损伤主要原因是什么", type: "radio", options: ["准备活动没做好", "技术动作错误", "太累/疲劳", "场地太滑/条件差", "其他"], condition: "Q5=是" },
  { id: 10, title: "您打球前做准备活动吗", type: "radio", options: ["每次都做", "经常做", "偶尔做", "很少/从不做"] },
  { id: 11, title: "您的准备活动一般做多久", type: "radio", options: ["5分钟以内", "5-10分钟", "10-20分钟", "20分钟以上"] },
  { id: 12, title: "您是否了解如何预防膝关节损伤", type: "radio", options: ["非常了解", "了解一些", "不太了解", "完全不了解"] },
  { id: 13, title: "受伤后您会怎么做", type: "radio", options: ["立即停止运动，休息/就医", "简单处理，继续打", "不管它，继续打", "没受过伤"] },
  { id: 14, title: "您对预防篮球膝关节损伤有什么建议", type: "text", options: [] }
];

// 基于现有数据的比例分布（从报告中提取的）
const distributions = {
  Q1: { "男": 0.71, "女": 0.29 },
  Q2: { "体育学院篮球专项": 0.21, "校篮球队成员": 0, "普通篮球爱好者（公体课/社团）": 0.79 },
  Q3: { "1年以下": 0.29, "1-3年": 0.43, "3-5年": 0, "5年以上": 0.28 },
  Q4: { "1-2次": 0.71, "3-4次": 0.21, "5次及以上": 0.08 },
  Q5: { "是": 0.57, "否": 0.43 },
  Q6: { "1次": 0.79, "2次": 0.07, "3次及以上": 0.14 }, // 条件：受过伤
  Q7: { 
    "半月板损伤": 0.43, 
    "髌骨劳损（膝盖骨疼痛）": 0.07, 
    "韧带损伤（内侧/前交叉）": 0.36, 
    "髌腱炎（跳跃膝）": 0.14, 
    "其他/不清楚": 0.21 
  }, // 多选，条件：受过伤
  Q8: { "起跳落地": 0.57, "急停变向": 0.14, "防守滑步": 0, "身体对抗": 0.07, "其他": 0.21 }, // 条件：受过伤
  Q9: { "准备活动没做好": 0.29, "技术动作错误": 0.21, "太累/疲劳": 0.29, "场地太滑/条件差": 0.14, "其他": 0.07 }, // 条件：受过伤
  Q10: { "每次都做": 0.29, "经常做": 0, "偶尔做": 0.71, "很少/从不做": 0 },
  Q11: { "5分钟以内": 0.43, "5-10分钟": 0.36, "10-20分钟": 0.14, "20分钟以上": 0.07 },
  Q12: { "非常了解": 0.29, "了解一些": 0.29, "不太了解": 0.36, "完全不了解": 0.07 },
  Q13: { "立即停止运动，休息/就医": 0.64, "简单处理，继续打": 0.21, "不管它，继续打": 0.07, "没受过伤": 0.07 }
};

// 随机选择函数（带权重）
function weightedRandom(options, weights) {
  const total = weights.reduce((a, b) => a + b, 0);
  let random = Math.random() * total;
  for (let i = 0; i < options.length; i++) {
    random -= weights[i];
    if (random <= 0) return options[i];
  }
  return options[options.length - 1];
}

// 简单随机选择
function randomChoice(options) {
  return options[Math.floor(Math.random() * options.length)];
}

// 多选随机选择（可选多个）
function randomMultiChoice(options, min = 1, max = 3) {
  const count = Math.floor(Math.random() * (max - min + 1)) + min;
  const shuffled = [...options].sort(() => 0.5 - Math.random());
  return shuffled.slice(0, count);
}

// 根据分布生成答案
function generateByDistribution(dist) {
  const options = Object.keys(dist);
  const weights = Object.values(dist);
  return weightedRandom(options, weights);
}

// 生成单条数据
function generateRecord(index) {
  const record = { "序号": index + 1 };
  
  // Q1: 性别
  record["Q1_性别"] = generateByDistribution(distributions.Q1);
  
  // Q2: 身份
  record["Q2_身份"] = generateByDistribution(distributions.Q2);
  
  // Q3: 运动年限
  record["Q3_运动年限"] = generateByDistribution(distributions.Q3);
  
  // Q4: 每周频率
  record["Q4_每周频率"] = generateByDistribution(distributions.Q4);
  
  // Q5: 是否受过伤
  record["Q5_是否受伤"] = generateByDistribution(distributions.Q5);
  
  // Q6-Q9: 只有受伤的人才回答
  if (record["Q5_是否受伤"] === "是") {
    record["Q6_受伤次数"] = generateByDistribution(distributions.Q6);
    record["Q7_损伤类型"] = randomMultiChoice(questions[6].options, 1, 3).join(";");
    record["Q8_损伤情况"] = generateByDistribution(distributions.Q8);
    record["Q9_损伤原因"] = generateByDistribution(distributions.Q9);
  } else {
    record["Q6_受伤次数"] = "";
    record["Q7_损伤类型"] = "";
    record["Q8_损伤情况"] = "";
    record["Q9_损伤原因"] = "";
  }
  
  // Q10: 准备活动
  record["Q10_准备活动"] = generateByDistribution(distributions.Q10);
  
  // Q11: 准备活动时长
  record["Q11_活动时长"] = generateByDistribution(distributions.Q11);
  
  // Q12: 预防知识了解
  record["Q12_预防知识"] = generateByDistribution(distributions.Q12);
  
  // Q13: 受伤后做法
  record["Q13_受伤后做法"] = generateByDistribution(distributions.Q13);
  
  // Q14: 建议（随机填充约30%）
  const suggestions = [
    "做好热身运动",
    "加强腿部肌肉训练",
    "使用护膝保护",
    "注意场地安全",
    "控制运动强度",
    "学习正确技术动作",
    "运动后做好拉伸",
    "穿戴合适的篮球鞋",
    "避免疲劳运动",
    "定期检查膝盖状况"
  ];
  record["Q14_建议"] = Math.random() > 0.7 ? randomChoice(suggestions) : "";
  
  return record;
}

// 生成CSV内容
function generateCSV(count) {
  const headers = [
    "序号", "Q1_性别", "Q2_身份", "Q3_运动年限", "Q4_每周频率", 
    "Q5_是否受伤", "Q6_受伤次数", "Q7_损伤类型", "Q8_损伤情况", "Q9_损伤原因",
    "Q10_准备活动", "Q11_活动时长", "Q12_预防知识", "Q13_受伤后做法", "Q14_建议"
  ];
  
  let csv = "\uFEFF" + headers.join(",") + "\n";
  
  for (let i = 0; i < count; i++) {
    const record = generateRecord(i);
    const row = headers.map(h => {
      const val = record[h];
      // 如果值包含逗号或引号，需要转义
      if (typeof val === 'string' && (val.includes(',') || val.includes('"'))) {
        return `"${val.replace(/"/g, '""')}"`;
      }
      return val;
    });
    csv += row.join(",") + "\n";
  }
  
  return csv;
}

// 主程序
const count = Math.floor(Math.random() * 21) + 50; // 50-70份
console.log(`正在生成 ${count} 份模拟数据...`);

const csvContent = generateCSV(count);
fs.writeFileSync('survey_data_模拟数据.csv', csvContent, 'utf8');

console.log(`✅ 成功生成 ${count} 份模拟数据，保存到 survey_data_模拟数据.csv`);

// 生成统计摘要
console.log("\n📊 数据分布预览:");
const records = [];
for (let i = 0; i < count; i++) {
  records.push(generateRecord(i));
}

// 统计各题分布
const stats = {};
records.forEach(r => {
  Object.keys(r).forEach(key => {
    if (key === "序号" || key === "Q14_建议") return;
    if (!stats[key]) stats[key] = {};
    const val = r[key] || "(未填)";
    stats[key][val] = (stats[key][val] || 0) + 1;
  });
});

// 打印统计
Object.keys(stats).forEach(key => {
  console.log(`\n${key}:`);
  Object.entries(stats[key]).forEach(([val, cnt]) => {
    const pct = ((cnt / count) * 100).toFixed(1);
    console.log(`  ${val}: ${cnt}人 (${pct}%)`);
  });
});
