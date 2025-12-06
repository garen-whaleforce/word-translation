/**
 * Job tracking and state management
 */

export type JobStatus =
  | "uploading"
  | "converting"
  | "parsing-docx"
  | "translating"
  | "qa-check"
  | "retranslating"
  | "packing"
  | "done"
  | "error"
  | "cancelled";

export interface JobUsage {
  prompt: number;
  completion: number;
  reasoning: number;
}

export interface JobState {
  id: string;
  fileName: string;
  status: JobStatus;
  progress: number; // 0-100
  stepMessage: string;
  createdAt: number;
  startedAt?: number;
  finishedAt?: number;
  errorMessage?: string;
  cancelled: boolean;
  abortController: AbortController;
  totalSegments: number;
  doneSegments: number;
  usage: JobUsage;
  costUSD?: number;
  outputPath?: string;
}

// In-memory job store
const jobs = new Map<string, JobState>();

/**
 * Create a new job with initial state
 */
export function createJob(id: string, fileName: string): JobState {
  const job: JobState = {
    id,
    fileName,
    status: "uploading",
    progress: 0,
    stepMessage: "準備中...",
    createdAt: Date.now(),
    cancelled: false,
    abortController: new AbortController(),
    totalSegments: 0,
    doneSegments: 0,
    usage: {
      prompt: 0,
      completion: 0,
      reasoning: 0,
    },
  };
  jobs.set(id, job);
  return job;
}

/**
 * Get a job by ID
 */
export function getJob(id: string): JobState | undefined {
  return jobs.get(id);
}

/**
 * Cancel a job - sets cancelled flag and aborts any ongoing operations
 */
export function cancelJob(id: string): boolean {
  const job = jobs.get(id);
  if (!job) return false;

  job.cancelled = true;
  job.status = "cancelled";
  job.stepMessage = "已取消";
  job.finishedAt = Date.now();
  job.abortController.abort();
  return true;
}

/**
 * Update job with partial data
 */
export function updateJob(job: JobState, patch: Partial<JobState>): void {
  Object.assign(job, patch);
}

/**
 * Mark job as finished and calculate cost
 */
export function finishJob(job: JobState, status: JobStatus = "done"): void {
  job.status = status;
  job.finishedAt = Date.now();
  job.progress = 100;

  // Calculate cost based on token usage (GPT-5-mini Data Zone pricing)
  // Input: $0.28 per 1M tokens, Output: $2.20 per 1M tokens
  const inputPricePerM = parseFloat(process.env.AZURE_OPENAI_INPUT_PRICE_PER_1M || "0.28");
  const outputPricePerM = parseFloat(process.env.AZURE_OPENAI_OUTPUT_PRICE_PER_1M || "2.20");

  const inputCost = (job.usage.prompt / 1_000_000) * inputPricePerM;
  const outputCost = ((job.usage.completion + job.usage.reasoning) / 1_000_000) * outputPricePerM;
  job.costUSD = inputCost + outputCost;
}

/**
 * Get elapsed time in seconds
 */
export function getElapsedSeconds(job: JobState): number {
  if (!job.startedAt) return 0;
  const endTime = job.finishedAt ?? Date.now();
  return Math.round((endTime - job.startedAt) / 1000);
}

/**
 * Delete a job from storage (for cleanup)
 */
export function deleteJob(id: string): boolean {
  return jobs.delete(id);
}

/**
 * Get all jobs (for debugging)
 */
export function getAllJobs(): JobState[] {
  return Array.from(jobs.values());
}
