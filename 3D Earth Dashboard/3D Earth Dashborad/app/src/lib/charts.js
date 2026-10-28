import { AXIS, CATS, REGIONS, css, draw } from './data.js';

/* 三张固定图 + 一张随下钻层级变的趋势图。配置都收拢在这一个模块里，
   好让组件在 effect 里按需重画。 */
export function trendOpt(level, curRegion, curCity) {
  const MON = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'];
  const scale = level === 0 ? 1 : level === 1 ? .22 : .05;
  let s2 = 90210 + (curRegion || '').length * 31 + (curCity || '').length * 7;
  const r = () => (s2 = (s2 * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff;
  const act = MON.map((_, i) => +(((1.05 + i * .07) * (0.85 + r() * .35)) * scale).toFixed(2));
  const tgt = MON.map((_, i) => +(((1.12 + i * .066)) * scale).toFixed(2));
  const cut = 8;
  const acc = css('--s1');
  return {
    animation: false,
    grid: {left: 56, right: 26, top: 40, bottom: 30},
    legend: {top: 4, right: 6, itemWidth: 14, itemHeight: 4, textStyle: {color: css('--muted'), fontSize: 14}},
    xAxis: {type: 'category', data: MON, boundaryGap: false, ...AXIS(), splitLine: {show: false}},
    yAxis: {type: 'value', name: '亿元', nameTextStyle: {color: css('--muted'), fontSize: 13}, ...AXIS()},
    series: [
      {name: '实际', type: 'line', smooth: .3, symbol: 'circle', symbolSize: 6,
       data: act.map((v, i) => i < cut ? v : null),
       lineStyle: {width: 3, color: acc, shadowColor: acc, shadowBlur: 14},
       itemStyle: {color: acc},
       areaStyle: {color: {type: 'linear', x: 0, y: 0, x2: 0, y2: 1, colorStops: [
         {offset: 0, color: 'color-mix(in srgb,' + acc + ' 45%,transparent)'},
         {offset: 1, color: 'color-mix(in srgb,' + acc + ' 0%,transparent)'}]}}},
      {name: '目标', type: 'line', smooth: .3, symbol: 'none', data: tgt,
       lineStyle: {width: 2, color: css('--s5'), type: 'dashed'}},
      {name: '预测', type: 'line', smooth: .3, symbol: 'none',
       data: act.map((v, i) => i >= cut - 1 ? v : null),
       lineStyle: {width: 2.4, color: acc, type: 'dotted', opacity: .6}}
    ]
  };
}

export function renderRest() {
  draw('mix', {
    animation: false,
    series: [{
      type: 'pie', radius: ['42%', '64%'], center: ['50%', '52%'],
      itemStyle: {borderColor: css('--panel'), borderWidth: 3},
      color: [css('--s1'), css('--s3'), css('--s2'), css('--s5'), css('--s4')],
      label: {color: css('--ink'), fontSize: 13.5, formatter: '{b}\n{d}%', lineHeight: 17,
              alignTo: 'edge', edgeDistance: 6},
      labelLine: {length: 8, length2: 10, maxSurfaceAngle: 80, lineStyle: {color: css('--line')}},
      data: CATS.map(c => ({name: c.n, value: c.v}))
    }]
  });

  const TURN = REGIONS.slice(0, 5);
  draw('turn', {
    animation: false,
    grid: {left: 12, right: 44, top: 10, bottom: 10, containLabel: true},
    xAxis: {type: 'value', ...AXIS(), axisLabel: {show: false}, splitLine: {show: false}},
    yAxis: {type: 'category', data: TURN.map(r => r.n.replace('大区', '')).reverse(), ...AXIS(),
            splitLine: {show: false}, axisLabel: {color: css('--ink'), fontSize: 14}},
    series: [{
      type: 'bar', barWidth: 13,
      data: TURN.map((r, i) => 32 + i * 6.4).reverse(),
      itemStyle: {borderRadius: [0, 3, 3, 0],
        color: p => p.value > 48 ? css('--s6') : p.value > 40 ? css('--s5') : css('--s2')},
      label: {show: true, position: 'right', formatter: '{c} 天', color: css('--muted'), fontSize: 14},
      markLine: {silent: true, symbol: 'none', data: [{xAxis: 45}],
        lineStyle: {color: css('--s6'), type: 'dashed', width: 1.5},
        label: {formatter: '红线 45', color: css('--s6'), fontSize: 12, position: 'end'}}
    }]
  });

  const MON6 = ['3月', '4月', '5月', '6月', '7月', '8月'];
  draw('cash', {
    animation: false,
    grid: {left: 54, right: 20, top: 34, bottom: 28},
    legend: {top: 2, right: 2, itemWidth: 12, itemHeight: 4, textStyle: {color: css('--muted'), fontSize: 13}},
    xAxis: {type: 'category', data: MON6, ...AXIS(), splitLine: {show: false}},
    yAxis: {type: 'value', name: '万元', nameTextStyle: {color: css('--muted'), fontSize: 13}, ...AXIS()},
    series: [
      {name: '应收', type: 'bar', barWidth: 12, data: [4200, 4650, 4380, 5120, 4980, 5340],
       itemStyle: {color: css('--faint'), borderRadius: [2, 2, 0, 0]}},
      {name: '实收', type: 'bar', barWidth: 12, data: [3680, 4020, 3910, 4410, 4180, 4610],
       itemStyle: {color: css('--s1'), borderRadius: [2, 2, 0, 0]}},
      {name: '经营现金流', type: 'line', smooth: .3, symbolSize: 6,
       data: [980, 1240, 1050, 1520, 1310, 1680],
       lineStyle: {width: 2.4, color: css('--s2'), shadowColor: css('--s2'), shadowBlur: 10},
       itemStyle: {color: css('--s2')}}
    ]
  });
}
