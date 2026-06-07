export function buildClientStatusMessage(message: string, sheetIsEmpty: boolean): string {
  const lower = message.toLowerCase();

  if (sheetIsEmpty) {
    if (/\b(create|generate|populate|add|fill|make|build|dummy|sample|row|header)\b/.test(lower)) {
      return 'Creating your table…';
    }
    return 'Working on your request…';
  }

  if (/\b(create|generate|populate)\b/.test(lower)) {
    return 'Updating your sheet…';
  }

  return 'Working on your request…';
}

export function isSimpleCreateTask(message: string, sheetIsEmpty: boolean): boolean {
  const lower = message.toLowerCase();
  return (
    sheetIsEmpty &&
    /\b(create|generate|add|populate|fill|make|build|give|dummy|sample)\b/.test(lower) &&
    /\b(row|rows|header|headers|column|columns|table)\b/.test(lower)
  );
}
