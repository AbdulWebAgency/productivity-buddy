// Public engine types.
export type EngineFile = { name: string; buffer: ArrayBuffer };

export type EngineResult = {
  buffer: Buffer;
  stats: Record<string, unknown>;
  warnings: string[];
};
