/** Limits that the code enforces and the help text quotes. Change one here, and both follow. */
export const MAX_FILES = 10;
export const MAX_UPLOAD_BYTES = 250 * 1024 * 1024;
export const MAX_DOWNLOAD_BYTES = 250 * 1024 * 1024;
export const MAX_IMAGES = 10;
export const MAX_IMAGE_BYTES = 25 * 1024 * 1024;
export const DEFAULT_TOP = 20;
export const MAX_TOP = 1000;
export const MAX_ALL = 20_000;
export const DEFAULT_SEARCH_COUNT = 25;
export const MAX_SEARCH_COUNT = 1000;
export const DEFAULT_CONTEXT = 5;
export const MAX_CONTEXT = 100;
/** The card router flow refuses a payload larger than this. */
export const MAX_CARD_BYTES = 256 * 1024;
/** Teams refuses a card larger than this: "The payload is too large. Please make sure the size is less than 28KB." */
export const MAX_TEAMS_CARD_BYTES = 28 * 1024;

/** Bytes as whole megabytes, for help and error text. */
export const mb = (bytes: number) => `${Math.round(bytes / 1024 / 1024)} MB`;
