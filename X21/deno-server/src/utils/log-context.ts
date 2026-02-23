let currentUserEmail: string | null = null;

export function setLogUserEmail(email: string | null): void {
  currentUserEmail = email ?? null;
}

export function getLogUserEmail(): string | null {
  return currentUserEmail;
}
