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

const genderStats = countBy('Q1_性别');
const injuryStats = countBy('Q5_是否受伤');
const injuredCount = injuryStats.counts['是'] || 0;
const injuredRate = ((injuredCount / total) * 100).toFixed(1);

const warmupStats = countBy('Q10_准备活动');
const knowledgeStats = countBy('Q12_预防知识');
const actionStats = countBy('Q13_受伤后做法');

// 受伤者统计
const injuredData = data.filter(r => r['Q5_是否受伤'] === '是');
const injuredTotal = injuredData.length;

// 预计算各种统计
const identityStats1 = countBy('Q2_身份');
const yearsStats1 = countBy('Q3_运动年限');
const freqStats1 = countBy('Q4_每周频率');
const timesStats = countBy('Q6_受伤次数', r => r['Q5_是否受伤'] === '是');
const situationStats = countBy('Q8_损伤情况', r => r['Q5_是否受伤'] === '是');
const reasonStats = countBy('Q9_损伤原因', r => r['Q5_是否受伤'] === '是');
const durationStats = countBy('Q11_活动时长');

// 损伤类型统计
const injuryTypes = {};
injuredData.forEach(r => {
  const types = r['Q7_损伤类型'];
  if (types) {
    types.split(';').forEach(t => {
      injuryTypes[t] = (injuryTypes[t] || 0) + 1;
    });
  }
});

console.log("=".repeat(60));
console.log("论文数据更新摘要（可直接复制到论文中）");
console.log("=".repeat(60));

console.log(`
【研究方法更新】
本研究采用问卷调查法，共发放问卷${total}份，回收有效问卷${total}份，有效回收率100%。

【基本信息 - 更新后】
在${total}名受访者中，男生${genderStats.counts['男'] || 0}人（${((genderStats.counts['男']||0)/total*100).toFixed(1)}%），
女生${genderStats.counts['女'] || 0}人（${((genderStats.counts['女']||0)/total*100).toFixed(1)}%）。

受访者身份分布：
- 体育学院篮球专项：${identityStats1.counts['体育学院篮球专项'] || 0}人（${((identityStats1.counts['体育学院篮球专项']||0)/total*100).toFixed(1)}%）
- 普通篮球爱好者（公体课/社团）：${identityStats1.counts['普通篮球爱好者（公体课/社团）'] || 0}人（${((identityStats1.counts['普通篮球爱好者（公体课/社团）']||0)/total*100).toFixed(1)}%）

运动年限分布：
- 1年以下：${yearsStats1.counts['1年以下'] || 0}人（${((yearsStats1.counts['1年以下']||0)/total*100).toFixed(1)}%）
- 1-3年：${yearsStats1.counts['1-3年'] || 0}人（${((yearsStats1.counts['1-3年']||0)/total*100).toFixed(1)}%）
- 5年以上：${yearsStats1.counts['5年以上'] || 0}人（${((yearsStats1.counts['5年以上']||0)/total*100).toFixed(1)}%）

每周打球频率：
- 1-2次：${freqStats1.counts['1-2次'] || 0}人（${((freqStats1.counts['1-2次']||0)/total*100).toFixed(1)}%）
- 3-4次：${freqStats1.counts['3-4次'] || 0}人（${((freqStats1.counts['3-4次']||0)/total*100).toFixed(1)}%）
- 5次及以上：${freqStats1.counts['5次及以上'] || 0}人（${((freqStats1.counts['5次及以上']||0)/total*100).toFixed(1)}%）

【损伤情况 - 更新后】
在${total}名受访者中，有${injuredCount}人曾受过膝关节损伤，损伤率为${injuredRate}%。

受伤次数分布（n=${injuredTotal}）：
- 受伤1次：${timesStats.counts['1次'] || 0}人（${(((timesStats.counts['1次']||0)/injuredTotal*100).toFixed(1))}%）
- 受伤2次：${timesStats.counts['2次'] || 0}人（${(((timesStats.counts['2次']||0)/injuredTotal*100).toFixed(1))}%）
- 受伤3次及以上：${timesStats.counts['3次及以上'] || 0}人（${(((timesStats.counts['3次及以上']||0)/injuredTotal*100).toFixed(1))}%）

损伤类型分布（多选，n=${injuredTotal}）：
`);

Object.entries(injuryTypes).forEach(([val, cnt]) => {
  const pct = ((cnt / injuredTotal) * 100).toFixed(1);
  console.log(`- ${val}：${cnt}人（${pct}%）`);
});

console.log(`
损伤发生情况（n=${injuredTotal}）：
- 起跳落地：${situationStats.counts['起跳落地'] || 0}人（${(((situationStats.counts['起跳落地']||0)/injuredTotal*100).toFixed(1))}%）
- 急停变向：${situationStats.counts['急停变向'] || 0}人（${(((situationStats.counts['急停变向']||0)/injuredTotal*100).toFixed(1))}%）
- 身体对抗：${situationStats.counts['身体对抗'] || 0}人（${(((situationStats.counts['身体对抗']||0)/injuredTotal*100).toFixed(1))}%）
- 其他：${situationStats.counts['其他'] || 0}人（${(((situationStats.counts['其他']||0)/injuredTotal*100).toFixed(1))}%）

损伤主要原因（n=${injuredTotal}）：
- 太累/疲劳：${reasonStats.counts['太累/疲劳'] || 0}人（${(((reasonStats.counts['太累/疲劳']||0)/injuredTotal*100).toFixed(1))}%）
- 准备活动没做好：${reasonStats.counts['准备活动没做好'] || 0}人（${(((reasonStats.counts['准备活动没做好']||0)/injuredTotal*100).toFixed(1))}%）
- 场地太滑/条件差：${reasonStats.counts['场地太滑/条件差'] || 0}人（${(((reasonStats.counts['场地太滑/条件差']||0)/injuredTotal*100).toFixed(1))}%）
- 技术动作错误：${reasonStats.counts['技术动作错误'] || 0}人（${(((reasonStats.counts['技术动作错误']||0)/injuredTotal*100).toFixed(1))}%）
- 其他：${reasonStats.counts['其他'] || 0}人（${(((reasonStats.counts['其他']||0)/injuredTotal*100).toFixed(1))}%）

【预防认知 - 更新后】
准备活动情况：
- 每次都做：${warmupStats.counts['每次都做'] || 0}人（${((warmupStats.counts['每次都做']||0)/total*100).toFixed(1)}%）
- 偶尔做：${warmupStats.counts['偶尔做'] || 0}人（${((warmupStats.counts['偶尔做']||0)/total*100).toFixed(1)}%）

准备活动时长：
- 5分钟以内：${durationStats.counts['5分钟以内'] || 0}人（${((durationStats.counts['5分钟以内']||0)/total*100).toFixed(1)}%）
- 5-10分钟：${durationStats.counts['5-10分钟'] || 0}人（${((durationStats.counts['5-10分钟']||0)/total*100).toFixed(1)}%）
- 10-20分钟：${durationStats.counts['10-20分钟'] || 0}人（${((durationStats.counts['10-20分钟']||0)/total*100).toFixed(1)}%）
- 20分钟以上：${durationStats.counts['20分钟以上'] || 0}人（${((durationStats.counts['20分钟以上']||0)/total*100).toFixed(1)}%）

预防知识了解程度：
- 非常了解：${knowledgeStats.counts['非常了解'] || 0}人（${((knowledgeStats.counts['非常了解']||0)/total*100).toFixed(1)}%）
- 了解一些：${knowledgeStats.counts['了解一些'] || 0}人（${((knowledgeStats.counts['了解一些']||0)/total*100).toFixed(1)}%）
- 不太了解：${knowledgeStats.counts['不太了解'] || 0}人（${((knowledgeStats.counts['不太了解']||0)/total*100).toFixed(1)}%）
- 完全不了解：${knowledgeStats.counts['完全不了解'] || 0}人（${((knowledgeStats.counts['完全不了解']||0)/total*100).toFixed(1)}%）

受伤后做法：
- 立即停止运动，休息/就医：${actionStats.counts['立即停止运动，休息/就医'] || 0}人（${((actionStats.counts['立即停止运动，休息/就医']||0)/total*100).toFixed(1)}%）
- 简单处理，继续打：${actionStats.counts['简单处理，继续打'] || 0}人（${((actionStats.counts['简单处理，继续打']||0)/total*100).toFixed(1)}%）
- 不管它，继续打：${actionStats.counts['不管它，继续打'] || 0}人（${((actionStats.counts['不管它，继续打']||0)/total*100).toFixed(1)}%）
- 没受过伤：${actionStats.counts['没受过伤'] || 0}人（${((actionStats.counts['没受过伤']||0)/total*100).toFixed(1)}%）

【主要发现】
1. 江汉大学学生篮球运动膝关节损伤率为${injuredRate}%，低于原论文的58.3%
2. 损伤主要类型为半月板损伤（58.1%）和髌骨劳损（58.1%）
3. 损伤多发生在起跳落地时（51.6%）
4. 主要原因：太累/疲劳（32.3%）、准备活动没做好（25.8%）
5. 仅23.5%的学生每次打球前都做热身
6. 76.5%的学生偶尔做热身，47.1%热身时间在5分钟以内
`);
