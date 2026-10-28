import { useEffect, useRef } from "react";
import * as echarts from "echarts";
import type { ECElementEvent, EChartsOption } from "echarts";

interface ChartProps {
  option: EChartsOption;
  className?: string;
  label: string;
  onClick?: (params: ECElementEvent) => void;
}

export function Chart({ option, className = "", label, onClick }: ChartProps) {
  const host = useRef<HTMLDivElement>(null);
  const instance = useRef<echarts.ECharts | null>(null);

  useEffect(() => {
    if (!host.current) return;
    const chart = echarts.init(host.current);
    instance.current = chart;
    const observer = new ResizeObserver(() => chart.resize());
    observer.observe(host.current);
    return () => {
      observer.disconnect();
      chart.dispose();
      instance.current = null;
    };
  }, []);

  useEffect(() => {
    instance.current?.setOption(option, true);
  }, [option]);

  useEffect(() => {
    const chart = instance.current;
    if (!chart || !onClick) return;
    chart.on("click", onClick);
    return () => { chart.off("click", onClick); };
  }, [onClick]);

  return <div ref={host} className={`chart ${className}`} role="img" aria-label={label} />;
}
