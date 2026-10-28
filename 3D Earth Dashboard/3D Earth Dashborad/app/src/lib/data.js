import * as echarts from 'echarts';

/* 这一屏的数据与图表配置：值都是写死的演示经营数字。 */
/* 提到最前面：下面的数据定义在模块求值时就会读设计令牌，声明在后面会撞上 TDZ。 */
const css = v => getComputedStyle(document.body).getPropertyValue(v).trim();

/* ── 数据 ─────────────────────────────────────────────────────────── */
let s = 5150921;
const rnd = () => (s = (s * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff;
const yi = n => (n / 1e8).toFixed(2);

const REGIONS = [
  {n: '华东大区', v: 3.42e8, q: 3.20e8, hub: '上海'},
  {n: '华南大区', v: 2.86e8, q: 2.90e8, hub: '深圳'},
  {n: '华北大区', v: 2.31e8, q: 2.10e8, hub: '北京'},
  {n: '西南大区', v: 1.74e8, q: 1.95e8, hub: '成都'},
  {n: '华中大区', v: 1.52e8, q: 1.40e8, hub: '武汉'},
  {n: '东北大区', v: 0.96e8, q: 1.20e8, hub: '沈阳'},
  {n: '西北大区', v: 0.78e8, q: 0.70e8, hub: '西安'}
];
const CITIES = {
  '华东大区': ['上海', '杭州', '南京', '苏州'],
  '华南大区': ['深圳', '广州', '香港'],
  '华北大区': ['北京', '天津'],
  '西南大区': ['成都', '重庆', '昆明'],
  '华中大区': ['武汉', '长沙'],
  '东北大区': ['沈阳', '哈尔滨'],
  '西北大区': ['西安']
};
const CATS = [
  {n: '智能终端', v: 5.12e8}, {n: '云服务', v: 3.44e8}, {n: '工业模组', v: 2.68e8},
  {n: '配件耗材', v: 1.52e8}, {n: '技术服务', v: 0.83e8}
];

/* 干线：地球上的弧线与右栏表格是同一份数据，点哪边都能定位到另一边 */
const ROUTES = [
  {from: '上海', to: '鹿特丹', v: 8.42, d: -1.2},
  {from: '深圳', to: '洛杉矶', v: 7.16, d: +3.4},
  {from: '上海', to: '纽约', v: 6.38, d: +1.1},
  {from: '广州', to: '迪拜', v: 4.92, d: +5.2},
  {from: '北京', to: '法兰克福', v: 4.35, d: -0.6},
  {from: '深圳', to: '新加坡', v: 3.88, d: +2.8},
  {from: '成都', to: '莫斯科', v: 2.64, d: -4.1, level: 'warn'},
  {from: '香港', to: '伦敦', v: 2.41, d: +0.9},
  {from: '西安', to: '阿姆斯特丹', v: 1.96, d: +7.3},
  {from: '上海', to: '悉尼', v: 1.72, d: -2.2}
];

const HUBS = [
  {city: '上海', v: 96}, {city: '深圳', v: 88}, {city: '北京', v: 74}, {city: '广州', v: 66},
  {city: '香港', v: 58}, {city: '新加坡', v: 62}, {city: '法兰克福', v: 54}, {city: '鹿特丹', v: 51},
  {city: '纽约', v: 70}, {city: '洛杉矶', v: 64}, {city: '迪拜', v: 47},
  {city: '莫斯科', v: 28, level: 'warn'}, {city: '伦敦', v: 44}, {city: '悉尼', v: 31},
  {city: '成都', v: 36}, {city: '西安', v: 24}, {city: '阿姆斯特丹', v: 39},
];

/* echarts 实例按宿主 id 存一份，缩放时要一起 resize。 */
export const charts = {};
const draw = (id, opt) => {
  charts[id] = charts[id] || echarts.init(document.getElementById(id));
  charts[id].setOption(opt, true);
};
/* A function, not a const: the colours come from CSS variables and must be
   re-read on every rebuild, or a theme change leaves the axes on the old palette. */
const AXIS = () => ({
  axisLine: {lineStyle: {color: css('--line')}},
  axisTick: {show: false},
  axisLabel: {color: css('--muted'), fontSize: 14},
  splitLine: {lineStyle: {color: css('--line-soft')}}
});

/* ── 告警（滚动，不依赖 hover） ───────────────────────────────────── */
const ALERTS = [
  ['09:42', '西南大区达成率 89.2%，连续 3 个月低于红线', 'a'],
  ['09:31', '成都→莫斯科干线在途量环比 −4.1%，清关滞留', 'a'],
  ['09:18', '恒信科技 应收 240 万逾期 63 天', 'a'],
  ['08:56', '深圳南山店 单日销售额创年内新高 84.2 万', 'b'],
  ['08:44', '云服务品类毛利率环比下滑 2.1pp', 'b'],
  ['08:20', '东北大区 3 家门店未提交日报', 'b'],
  ['07:58', '工业模组 A2 批次退货率 4.8%，超阈值', 'a'],
  ['07:31', '集团经营现金流转正，环比 +28.2%', 'b']
];

/* ── 无人操作自动轮播，有操作暂停 ─────────────────────────────────── */
const TOUR = [
  {level: 0, region: null, city: null, hub: '上海'},
  {level: 1, region: '华东大区', city: null, hub: '上海'},
  {level: 1, region: '华南大区', city: null, hub: '深圳'},
  {level: 1, region: '西南大区', city: null, hub: '成都'}
];

export { css, rnd, yi, REGIONS, CITIES, CATS, ROUTES, HUBS, draw, AXIS, ALERTS, TOUR };
