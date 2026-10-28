export const usd = (value: number, digits = 2) => {
  const absolute = Math.abs(value);
  if (absolute >= 1_000_000_000) return `$${(value / 1_000_000_000).toFixed(digits)}B`;
  if (absolute >= 1_000_000) return `$${(value / 1_000_000).toFixed(digits)}M`;
  if (absolute >= 1_000) return `$${(value / 1_000).toFixed(digits)}K`;
  return `$${value.toFixed(0)}`;
};

export const compact = (value: number) => new Intl.NumberFormat("zh-TW", {
  notation: "compact",
  maximumFractionDigits: 1,
}).format(value);

export const integer = (value: number) => new Intl.NumberFormat("zh-TW", {
  maximumFractionDigits: 0,
}).format(value);

export const decimal = (value: number) => new Intl.NumberFormat("zh-TW", {
  maximumFractionDigits: 2,
}).format(value);

export const monthLabel = (month: string) => {
  const [year, value] = month.split("-");
  return `${year.slice(2)}.${value}`;
};

export const csvCell = (value: unknown) => {
  const text = value === null || value === undefined ? "" : String(value);
  return /[",\r\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
};
