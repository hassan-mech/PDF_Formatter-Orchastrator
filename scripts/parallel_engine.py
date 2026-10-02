"""
parallel_engine.py — High-performance parallel PDF page processing engine

Splits any PDF into page-chunks and dispatches them to a pool of workers.
Each worker runs independently — CPU-bound work (rendering, OCR, text extraction)
is done via ProcessPoolExecutor; I/O-bound work via ThreadPoolExecutor.

Architecture:
    ┌──────────────┐
    │   PDF Input  │
    └──────┬───────┘
           │ split into N chunks
    ┌──────▼───────────────────────────────┐
    │        ChunkDispatcher               │
    │  chunk_0 → Worker-0 (Process)        │
    │  chunk_1 → Worker-1 (Process)        │
    │  chunk_2 → Worker-2 (Process)        │
    │  ...                                 │
    └──────┬───────────────────────────────┘
           │ collect PageResult objects
    ┌──────▼───────┐
    │  ResultMerger│  → final ordered list
    └──────────────┘

Usage:
    from parallel_engine import ParallelPDFEngine, PageTask, WorkerMode

    engine = ParallelPDFEngine(pdf_path, workers=8, chunk_size=10)
    results = engine.run(WorkerMode.EXTRACT_TEXT)
"""

from __future__ import annotations

import os
import time
import traceback
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from enum import Enum, auto
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

try:
    import fitz  # PyMuPDF
except ImportError:
    raise ImportError("PyMuPDF is required: pip install pymupdf")


# ────────────────────────────────────────────────────────────────────────────
# Enums & Data Structures
# ────────────────────────────────────────────────────────────────────────────

class WorkerMode(Enum):
    INSPECT     = auto()   # summary info per page
    EXTRACT_TEXT = auto()  # extract selectable text
    EXTRACT_IMAGES = auto()  # render page as PNG bytes
    EXTRACT_TABLES = auto()  # detect & extract tables
    OCR         = auto()   # image-only page OCR
    FULL        = auto()   # all of the above


@dataclass
class PageTask:
    """Describes a single page task sent to a worker."""
    pdf_path: str          # serialisable path (str, not Path)
    page_index: int        # 0-based
    dpi: int = 150
    mode: str = "inspect"  # string form of WorkerMode name


@dataclass
class PageResult:
    """Result produced by a worker for a single page."""
    page_index: int
    page_number: int       # 1-based for display
    success: bool
    mode: str
    # Payload — populated depending on mode
    text: str = ""
    image_bytes: bytes = field(default=b"", repr=False)
    tables: List[List[List[str]]] = field(default_factory=list)
    is_image_only: bool = False
    width_pt: float = 0.0
    height_pt: float = 0.0
    image_count: int = 0
    drawing_count: int = 0
    elapsed_ms: float = 0.0
    error: str = ""


@dataclass
class ChunkResult:
    """Aggregated results for a chunk of pages."""
    chunk_id: int
    page_results: List[PageResult]
    elapsed_ms: float
    worker_pid: int


# ────────────────────────────────────────────────────────────────────────────
# Worker functions (module-level so they are picklable for multiprocessing)
# ────────────────────────────────────────────────────────────────────────────

def _process_page(task: PageTask) -> PageResult:
    """Process a single page — runs inside a worker process."""
    t0 = time.perf_counter()
    result = PageResult(
        page_index=task.page_index,
        page_number=task.page_index + 1,
        success=False,
        mode=task.mode,
    )
    try:
        import fitz  # reimport in worker process
        doc = fitz.open(task.pdf_path)
        page = doc[task.page_index]

        raw_text = page.get_text("text").strip()
        result.is_image_only = len(raw_text) == 0
        result.width_pt = page.rect.width
        result.height_pt = page.rect.height
        result.image_count = len(page.get_images(full=True))
        result.drawing_count = len(page.get_drawings())

        mode = task.mode.upper()

        # ── Text extraction ──────────────────────────────────────────────
        if mode in ("EXTRACT_TEXT", "FULL", "INSPECT"):
            result.text = raw_text

        # ── Image rendering ──────────────────────────────────────────────
        if mode in ("EXTRACT_IMAGES", "FULL"):
            mat = fitz.Matrix(task.dpi / 72, task.dpi / 72)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            result.image_bytes = pix.tobytes("png")

        # ── Table extraction ─────────────────────────────────────────────
        if mode in ("EXTRACT_TABLES", "FULL"):
            tabs = page.find_tables()
            result.tables = [
                [
                    [cell.strip().replace("\n", " ") if cell else "" for cell in row]
                    for row in tab.extract()
                ]
                for tab in tabs.tables
            ]

        doc.close()
        result.success = True

    except Exception as exc:
        result.error = traceback.format_exc()

    result.elapsed_ms = (time.perf_counter() - t0) * 1000
    return result


def _process_chunk(chunk_id: int, tasks: List[PageTask]) -> ChunkResult:
    """Process a list of page tasks sequentially inside one worker process."""
    t0 = time.perf_counter()
    page_results = [_process_page(t) for t in tasks]
    return ChunkResult(
        chunk_id=chunk_id,
        page_results=page_results,
        elapsed_ms=(time.perf_counter() - t0) * 1000,
        worker_pid=os.getpid(),
    )


# ────────────────────────────────────────────────────────────────────────────
# Chunk dispatcher
# ────────────────────────────────────────────────────────────────────────────

class ChunkDispatcher:
    """Splits page tasks into chunks and dispatches to a worker pool."""

    def __init__(
        self,
        tasks: List[PageTask],
        workers: int,
        chunk_size: int,
        use_threads: bool = False,
    ):
        self.tasks = tasks
        self.workers = workers
        self.chunk_size = chunk_size
        self.use_threads = use_threads  # ThreadPool for I/O, ProcessPool for CPU

    def _make_chunks(self) -> List[List[PageTask]]:
        chunks = []
        for i in range(0, len(self.tasks), self.chunk_size):
            chunks.append(self.tasks[i : i + self.chunk_size])
        return chunks

    def dispatch(self, progress_cb: Optional[Callable] = None) -> List[ChunkResult]:
        chunks = self._make_chunks()
        results: List[Optional[ChunkResult]] = [None] * len(chunks)

        Executor = ThreadPoolExecutor if self.use_threads else ProcessPoolExecutor

        with Executor(max_workers=self.workers) as pool:
            futures = {
                pool.submit(_process_chunk, cid, chunk): cid
                for cid, chunk in enumerate(chunks)
            }
            for future in as_completed(futures):
                cid = futures[future]
                try:
                    chunk_result = future.result()
                    results[cid] = chunk_result
                    if progress_cb:
                        done = sum(1 for r in results if r is not None)
                        progress_cb(done, len(chunks), chunk_result)
                except Exception as exc:
                    # Build a failed ChunkResult
                    results[cid] = ChunkResult(
                        chunk_id=cid,
                        page_results=[
                            PageResult(
                                page_index=t.page_index,
                                page_number=t.page_index + 1,
                                success=False,
                                mode=t.mode,
                                error=str(exc),
                            )
                            for t in chunks[cid]
                        ],
                        elapsed_ms=0.0,
                        worker_pid=-1,
                    )

        return [r for r in results if r is not None]


# ────────────────────────────────────────────────────────────────────────────
# Result Merger
# ────────────────────────────────────────────────────────────────────────────

class ResultMerger:
    """Merges chunk results into a single ordered list of PageResults."""

    @staticmethod
    def merge(chunk_results: List[ChunkResult]) -> List[PageResult]:
        all_pages: List[PageResult] = []
        for cr in chunk_results:
            all_pages.extend(cr.page_results)
        # Sort by original page index to restore order
        all_pages.sort(key=lambda r: r.page_index)
        return all_pages

    @staticmethod
    def summary(page_results: List[PageResult]) -> Dict[str, Any]:
        total = len(page_results)
        succeeded = sum(1 for r in page_results if r.success)
        image_only = sum(1 for r in page_results if r.is_image_only)
        failed = total - succeeded
        avg_ms = sum(r.elapsed_ms for r in page_results) / max(total, 1)
        return {
            "total_pages": total,
            "succeeded": succeeded,
            "failed": failed,
            "image_only_pages": image_only,
            "avg_ms_per_page": round(avg_ms, 2),
            "errors": [
                {"page": r.page_number, "error": r.error}
                for r in page_results
                if not r.success
            ],
        }


# ────────────────────────────────────────────────────────────────────────────
# High-level Engine API
# ────────────────────────────────────────────────────────────────────────────

class ParallelPDFEngine:
    """
    High-performance parallel PDF processing engine.

    Example
    -------
    engine = ParallelPDFEngine("input/large.pdf", workers=8, chunk_size=10)
    results = engine.run(WorkerMode.EXTRACT_TEXT)
    for r in results:
        print(r.page_number, r.text[:80])
    """

    def __init__(
        self,
        pdf_path: str | Path,
        workers: Optional[int] = None,
        chunk_size: int = 8,
        dpi: int = 150,
        pages: Optional[str] = None,  # e.g. '1-20' or None for all
    ):
        self.pdf_path = Path(pdf_path)
        if not self.pdf_path.exists():
            raise FileNotFoundError(f"PDF not found: {self.pdf_path}")

        # Auto-size workers: leave 2 cores free for OS + main process
        cpu = os.cpu_count() or 4
        self.workers = workers or max(1, cpu - 2)
        self.chunk_size = chunk_size
        self.dpi = dpi

        # Open briefly to get page count
        _doc = fitz.open(str(self.pdf_path))
        self.total_pages = _doc.page_count
        _doc.close()

        self.page_indices = self._parse_pages(pages)

    def _parse_pages(self, pages_str: Optional[str]) -> List[int]:
        if pages_str is None or pages_str.strip().lower() == "all":
            return list(range(self.total_pages))
        indices = []
        for part in pages_str.split(","):
            part = part.strip()
            if "-" in part:
                s, e = part.split("-", 1)
                indices.extend(range(int(s) - 1, int(e)))
            else:
                indices.append(int(part) - 1)
        return [i for i in indices if 0 <= i < self.total_pages]

    def _progress_cb(self, done: int, total: int, chunk: ChunkResult):
        pct = done / total * 100
        pages_done = sum(len(cr.page_results) for cr in [chunk])
        bar_len = 30
        filled = int(bar_len * done / total)
        bar = "█" * filled + "░" * (bar_len - filled)
        print(
            f"\r  [{bar}] {pct:5.1f}%  chunk {chunk.chunk_id+1}/{total}"
            f"  pid={chunk.worker_pid}  {chunk.elapsed_ms:.0f}ms",
            end="",
            flush=True,
        )

    def run(
        self,
        mode: WorkerMode = WorkerMode.INSPECT,
        verbose: bool = True,
    ) -> List[PageResult]:
        """Run the parallel engine and return ordered PageResults."""
        t_start = time.perf_counter()

        tasks = [
            PageTask(
                pdf_path=str(self.pdf_path),
                page_index=i,
                dpi=self.dpi,
                mode=mode.name,
            )
            for i in self.page_indices
        ]

        effective_workers = min(self.workers, len(tasks))
        effective_chunk = max(1, self.chunk_size)

        if verbose:
            print(f"\n⚡ ParallelPDFEngine starting")
            print(f"   PDF        : {self.pdf_path.name}")
            print(f"   Pages      : {len(self.page_indices)} / {self.total_pages}")
            print(f"   Mode       : {mode.name}")
            print(f"   Workers    : {effective_workers}  (of {os.cpu_count()} CPUs)")
            print(f"   Chunk size : {effective_chunk} pages/chunk")
            print(f"   Chunks     : {-(-len(tasks) // effective_chunk)}")  # ceil div

        dispatcher = ChunkDispatcher(
            tasks=tasks,
            workers=effective_workers,
            chunk_size=effective_chunk,
            use_threads=False,  # CPU-bound → ProcessPool
        )

        chunk_results = dispatcher.dispatch(
            progress_cb=self._progress_cb if verbose else None
        )
        if verbose:
            print()  # newline after progress bar

        page_results = ResultMerger.merge(chunk_results)
        summary = ResultMerger.summary(page_results)

        elapsed = time.perf_counter() - t_start

        if verbose:
            print(f"\n✅ Done in {elapsed:.2f}s")
            print(f"   Pages succeeded : {summary['succeeded']} / {summary['total_pages']}")
            print(f"   Image-only pages: {summary['image_only_pages']}  ← need OCR")
            print(f"   Avg per page    : {summary['avg_ms_per_page']} ms")
            if summary["failed"]:
                print(f"   ⚠️  Failed        : {summary['failed']}")
                for e in summary["errors"]:
                    print(f"      Page {e['page']}: {e['error'][:120]}")

        return page_results


# ────────────────────────────────────────────────────────────────────────────
# CLI
# ────────────────────────────────────────────────────────────────────────────

def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="parallel_engine.py — parallel PDF page processor",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("-i", "--input", required=True, help="Path to PDF file")
    parser.add_argument(
        "--mode",
        choices=[m.name.lower() for m in WorkerMode],
        default="inspect",
        help="Processing mode (default: inspect)",
    )
    parser.add_argument("--workers", type=int, default=None, help="Worker processes (default: cpu-2)")
    parser.add_argument("--chunk-size", type=int, default=8, help="Pages per chunk (default: 8)")
    parser.add_argument("--dpi", type=int, default=150, help="Render DPI for images (default: 150)")
    parser.add_argument("--pages", default=None, help="Page range e.g. '1-20' (default: all)")
    parser.add_argument("--output", "-o", default=None, help="Save text summary to file")
    args = parser.parse_args()

    mode = WorkerMode[args.mode.upper()]

    engine = ParallelPDFEngine(
        pdf_path=args.input,
        workers=args.workers,
        chunk_size=args.chunk_size,
        dpi=args.dpi,
        pages=args.pages,
    )

    results = engine.run(mode=mode, verbose=True)

    if args.output:
        out = Path(args.output)
        out.parent.mkdir(parents=True, exist_ok=True)
        with out.open("w", encoding="utf-8") as f:
            for r in results:
                f.write(f"=== Page {r.page_number} ===\n")
                if r.text:
                    f.write(r.text + "\n")
                if r.error:
                    f.write(f"[ERROR] {r.error}\n")
                f.write("\n")
        print(f"📄 Text saved to: {args.output}")


if __name__ == "__main__":
    main()
