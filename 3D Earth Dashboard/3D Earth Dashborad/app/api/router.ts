import { createRouter, publicQuery } from "./middleware";
import { procurementRouter } from "./procurementRouter";

export const appRouter = createRouter({
  ping: publicQuery.query(() => ({ ok: true, ts: Date.now() })),
  procurement: procurementRouter,
});

export type AppRouter = typeof appRouter;
