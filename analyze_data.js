const fs = require('fs');

// 读取CSV数据
const csvContent = fs.readFileSync('survey_data_模拟数据.csv', 'utf8');
const lines = csvContent.trim().split('\n');
const headers = lines[0].replace(/^\uFEFF/, '').split(',');
const data = [];

for (let i = 1; i < lines.length; i++) {
  const values = lines[i].split(',');
  const record = {};
  headers.forEach((h, idx) => {
    record[h] = values[idx] || '';
  });
  data.push(record);
}

const total = data.length;
console.log(`总样本量: ${total}人\n`);

// 统计函数
function countBy(field, filterFn = null) {
  const counts = {};
  const filtered = filterFn ? data.filter(filterFn) : data;
  filtered.forEach(r => {
    const val = r[field] || '(未填)';
    counts[val] = (counts[val] || 0) + 1;
  });
  return { counts, total: filtered.length };
}

function printStats(title, field, filterFn = null) {
  console.log(`\n=== ${title} ===`);
  const { counts, total: subTotal } = countBy(field, filterFn);
  Object.entries(counts).forEach(([val, cnt]) => {
    const pct = ((cnt / subTotal) * 100).toFixed(1);
    console.log(`  ${val}: ${cnt}人 (${pct}%)`);
  });
  console.log(`  小计: ${subTotal}人`);
  return counts;
}

// 1. 基本信息统计
console.log("=".repeat(50));
console.log("第一部分：基本信息");
console.log("=".repeat(50));

const genderStats = printStats("性别分布", "Q1_性别");
const identityStats = printStats("身份分布", "Q2_身份");
const yearsStats = printStats("运动年限", "Q3_运动年限");
const freqStats = printStats("每周频率", "Q4_每周频率");

// 2. 损伤情况统计
console.log("\n" + "=".repeat(50));
console.log("第二部分：损伤情况");
console.log("=".repeat(50));

const injuryStats = printStats("是否受过膝关节损伤", "Q5_是否受伤");
const injuredCount = injuryStats['是'] || 0;
const injuredRate = ((injuredCount / total) * 100).toFixed(1);
console.log(`\n损伤率: ${injuredRate}% (${injuredCount}/${total})`);

// 只统计受伤的人
const injuredData = data.filter(r => r['Q5_是否受伤'] === '是');
const injuredTotal = injuredData.length;

if (injuredTotal > 0) {
  printStats("受伤次数", "Q6_受伤次数", r => r['Q5_是否受伤'] === '是');
  
  // 损伤类型统计（多选）
  console.log("\n=== 损伤类型（多选）===");
  const injuryTypes = {};
  injuredData.forEach(r => {
    const types = r['Q7_损伤类型'];
    if (types) {
      types.split(';').forEach(t => {
        injuryTypes[t] = (injuryTypes[t] || 0) + 1;
      });
    }
  });
  Object.entries(injuryTypes).forEach(([val, cnt]) => {
    const pct = ((cnt / injuredTotal) * 100).toFixed(1);
    console.log(`  ${val}: ${cnt}人 (${pct}%)`);
  });
  
  printStats("损伤发生情况", "Q8_损伤情况", r => r['Q5_是否受伤'] === '是');
  printStats("损伤主要原因", "Q9_损伤原因", r => r['Q5_是否受伤'] === '是');
}

// 3. 预防认知统计
console.log("\n" + "=".repeat(50));
console.log("第三部分：预防认知");
console.log("=".repeat(50));

printStats("准备活动频率", "Q10_准备活动");
printStats("准备活动时长", "Q11_活动时长");
printStats("预防知识了解程度", "Q12_预防知识");
printStats("受伤后做法", "Q13_受伤后做法");

// 4. 交叉分析
console.log("\n" + "=".repeat(50));
console.log("第四部分：交叉分析");
console.log("=".repeat(50));

// 性别与损伤的关系
console.log("\n=== 性别与损伤关系 ===");
const maleData = data.filter(r => r['Q1_性别'] === '男');
const femaleData = data.filter(r => r['Q1_性别'] === '女');
const maleInjured = maleData.filter(r => r['Q5_是否受伤'] === '是').length;
const femaleInjured = femaleData.filter(r => r['Q5_是否受伤'] === '是').length;
console.log(`  男性: ${maleData.length}人，受伤${maleInjured}人 (${((maleInjured/maleData.length)*100).toFixed(1)}%)`);
console.log(`  女性: ${femaleData.length}人，受伤${femaleInjured}人 (${((femaleInjured/femaleData.length)*100).toFixed(1)}%)`);

// 运动年限与损伤的关系
console.log("\n=== 运动年限与损伤关系 ===");
const yearsGroups = ['1年以下', '1-3年', '3-5年', '5年以上'];
yearsGroups.forEach(y => {
  const group = data.filter(r => r['Q3_运动年限'] === y);
  const injured = group.filter(r => r['Q5_是否受伤'] === '是').length;
  if (group.length > 0) {
    console.log(`  ${y}: ${group.length}人，受伤${injured}人 (${((injured/group.length)*100).toFixed(1)}%)`);
  }
});

// 准备活动与损伤的关系
console.log("\n=== 准备活动与损伤关系 ===");
const warmupGroups = ['每次都做', '经常做', '偶尔做', '很少/从不做'];
warmupGroups.forEach(w => {
  const group = data.filter(r => r['Q10_准备活动'] === w);
  const injured = group.filter(r => r['Q5_是否受伤'] === '是').length;
  if (group.length > 0) {
    console.log(`  ${w}: ${group.length}人，受伤${injured}人 (${((injured/group.length)*100).toFixed(1)}%)`);
  }
});

// 生成论文数据摘要
console.log("\n" + "=".repeat(50));
console.log("论文数据摘要（可直接复制到论文中）");
console.log("=".repeat(50));

console.log(`
【基本信息】
本研究共收集有效问卷${total}份。其中男生${genderStats['男'] || 0}人（${((genderStats['男']||0)/total*100).toFixed(1)}%），女生${genderStats['女'] || 0}人（${((genderStats['女']||0)/total*100).toFixed(1)}%）。

【损伤情况】
在${total}名受访者中，有${injuredCount}人曾受过膝关节损伤，损伤率为${injuredRate}%。

受伤次数分布：
- 受伤1次：${countBy('Q6_受伤次数', r => r['Q5_是否受伤'] === '是').counts['1次'] || 0}人
- 受伤2次：${countBy('Q6_受伤次数', r => r['Q5_是否受伤'] === '是').counts['2次'] || 0}人  
- 受伤3次及以上：${countBy('Q6_受伤次数', r => r['Q5_是否受伤'] === '是').counts['3次及以上'] || 0}人

【预防认知】
- 每次打球前都做热身：${(identityStats = countBy('Q10_准备活动').counts['每次都做'] || 0)}人（${((identityStats)/total*100).toFixed(1)}%）
- 偶尔做热身：${(identityStats = countBy('Q10_准备活动').counts['偶尔做'] || 0)}人（${((identityStats)/total*100).toFixed(1)}%）

【受伤后处理】
- 立即停止运动，休息/就医：${countBy('Q13_受伤后做法').counts['立即停止运动，休息/就医'] || 0}人
- 简单处理，继续打：${countBy('Q13_受伤后做法').counts['简单处理，继续打'] || 0}人
- 不管它，继续打：${countBy('Q13_受伤后做法').counts['不管它，继续打'] || 0}人
`);
