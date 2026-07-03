import { createServerFn } from "@tanstack/react-start";
import { z } from "zod";

const archiveSchema = z.object({
  filename: z.string().min(1).max(255),
  base64: z.string().min(1),
  meta: z.object({
    serverCount: z.number().int().nonnegative(),
    companyName: z.string().optional(),
    vendorName: z.string().optional(),
    periode: z.string().optional(),
    hosts: z
      .array(
        z.object({
          hostname: z.string(),
          ipAddress: z.string().optional(),
          osRelease: z.string().optional(),
        }),
      )
      .default([]),
  }),
});

export const archiveExport = createServerFn({ method: "POST" })
  .inputValidator((d) => archiveSchema.parse(d))
  .handler(async ({ data }) => {
    const { archiveDocx } = await import("./exports.server");
    return archiveDocx(data);
  });

export const fetchExportHistory = createServerFn({ method: "GET" })
  .handler(async () => {
    const { listExports } = await import("./exports.server");
    return listExports(50);
  });
