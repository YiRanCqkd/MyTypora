export function toggleTaskLine(line: string) {
  const pattern = /^(\s*(?:>\s*)*(?:[-+*]|\d+[.)])\s+\[)([ xX])(\]\s*)/;
  if (!pattern.test(line)) return null;
  return line.replace(pattern, (_match, prefix, marker, suffix) => {
    return `${prefix}${marker === ' ' ? 'x' : ' '}${suffix}`;
  });
}
