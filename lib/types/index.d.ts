/**
 * dsh-office-tools host plugin.
 *
 * Registers seven model-facing tools on `ctx.tools`:
 *
 *   word_create / word_read
 *   excel_create / excel_read / excel_update
 *   ppt_create / ppt_read
 *
 * All file access is confined to the calling agent's session workspace and
 * every registration is wrapped in `ctx.effect` so Cordis disposes the tools
 * with the plugin fiber.
 *
 * The PowerPoint pair is config-gated: dedicated presentation plugins such as
 * dsh-ppt register a colliding `ppt_create`, and DSH refuses duplicate tool
 * names at startup, so profiles running one of those set `enablePptTools: false`
 * to load this plugin for Word/Excel only.
 */
import type { Context } from '@deepseek-ai/cordis';
import z from '@deepseek-ai/schemastery';
/** Plugin identity for cordis.yml rows. */
export declare const name = "dsh-office-tools";
/** The tool registry is the only runtime service this plugin requires. */
export declare const inject: string[];
/** Host plugin configuration, validated at load by the Loader. */
export interface Config {
    /** Register `ppt_create` / `ppt_read`. */
    enablePptTools: boolean;
}
/** Configuration schema; the callable form applies schema defaults. */
export declare const Config: z<Schemastery.ObjectS<{
    enablePptTools: z<boolean, boolean>;
}>, Schemastery.ObjectT<{
    enablePptTools: z<boolean, boolean>;
}>>;
export declare function apply(ctx: Context, config?: Config): void;
