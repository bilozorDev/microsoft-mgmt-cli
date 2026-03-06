export function escapePS(value: string): string {
  return value.replace(/'/g, "''");
}
