/* global CustomFunctions */
import { askHermes } from "../shared/hermes";

function ikey(name, args) {
  return name + ":" + JSON.stringify(args);
}
function clean(s) {
  return String(s).trim().replace(/^["'`]+|["'`]+$/g, "").trim();
}

/**
 * Classify a value with Hermes.
 * @customfunction CLASSIFY
 * @param {string} value Text to classify.
 * @param {string} instruction How to classify, e.g. "lead quality: hot/warm/cold".
 * @returns {Promise<string>} The label.
 */
export async function classify(value, instruction) {
  const out = await askHermes(
    [{ role: "user", content: `Classify the text per this instruction: ${instruction}.\nReply with ONLY the label, no punctuation.\n\nText: ${value}` }],
    { idempotencyKey: ikey("CLASSIFY", [value, instruction]) }
  );
  return clean(out);
}

/**
 * Extract a field from text with Hermes.
 * @customfunction EXTRACT
 * @param {string} value Source text.
 * @param {string} what What to extract, e.g. "company name".
 * @returns {Promise<string>} The extracted value.
 */
export async function extract(value, what) {
  const out = await askHermes(
    [{ role: "user", content: `Extract the ${what} from the text. Reply with ONLY the value, or an empty string if none.\n\nText: ${value}` }],
    { idempotencyKey: ikey("EXTRACT", [value, what]) }
  );
  return clean(out);
}

/**
 * Summarize a range with Hermes.
 * @customfunction SUMMARIZE
 * @param {any[][]} values Range of values.
 * @returns {Promise<string>} One-sentence summary.
 */
export async function summarize(values) {
  const out = await askHermes(
    [{ role: "user", content: `Summarize this data in ONE short sentence:\n${JSON.stringify(values)}` }],
    { idempotencyKey: ikey("SUMMARIZE", values) }
  );
  return clean(out);
}

/**
 * Get an Excel formula from a description with Hermes.
 * @customfunction FORMULA_HELP
 * @param {string} goal What you want the formula to do.
 * @returns {Promise<string>} A single Excel formula.
 */
export async function formulaHelp(goal) {
  const out = await askHermes(
    [{ role: "user", content: `Give a single Excel formula that accomplishes: ${goal}. Reply with ONLY the formula, starting with =.` }],
    { idempotencyKey: ikey("FORMULA_HELP", [goal]) }
  );
  return clean(out);
}
