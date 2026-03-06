import { randomInt } from "crypto";

const UPPERCASE = "ABCDEFGHJKLMNPQRSTUVWXYZ";
const LOWERCASE = "abcdefghjkmnpqrstuvwxyz";
const DIGITS = "23456789";
const SYMBOLS = "!@#$%^&*-_=+?";

const ALL_CHARS = UPPERCASE + LOWERCASE + DIGITS + SYMBOLS;

export function generatePassword(length = 16): string {
  // Guarantee at least one from each class
  const required = [
    pick(UPPERCASE),
    pick(LOWERCASE),
    pick(DIGITS),
    pick(SYMBOLS),
  ];

  const rest = Array.from({ length: length - required.length }, () => pick(ALL_CHARS));
  const chars = [...required, ...rest];

  // Fisher-Yates shuffle with crypto random
  for (let i = chars.length - 1; i > 0; i--) {
    const j = cryptoRandomInt(i + 1);
    const tmp = chars[i]!;
    chars[i] = chars[j]!;
    chars[j] = tmp;
  }

  return chars.join("");
}

function pick(charset: string): string {
  return charset[cryptoRandomInt(charset.length)]!;
}

function cryptoRandomInt(max: number): number {
  return randomInt(max);
}

export function validatePassword(pw: string): string | undefined {
  if (pw.length < 8) return "Must be at least 8 characters";
  if (!/[A-Z]/.test(pw)) return "Must contain at least one uppercase letter";
  if (!/[0-9]/.test(pw)) return "Must contain at least one number";
  if (!/[^a-zA-Z0-9]/.test(pw)) return "Must contain at least one symbol";
}
