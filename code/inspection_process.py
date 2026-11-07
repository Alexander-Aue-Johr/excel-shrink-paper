"""Reuse a bounded parser process while retaining a deadline for each object."""

import argparse
from contextlib import redirect_stdout
import json
from pathlib import Path
from queue import Empty, Queue
import subprocess
import sys
from threading import Thread
import uuid

RESULT_PREFIX = b'CORPUS_INSPECTION_RESULT '


class InspectorProcess:
    """Own one lazily started parser and replace it after timeout or bounded use.

    Only the controller claims objects. The child uses the existing
    ``workers.inspect_one`` entry point and closes its Ledger after each object.
    A response is published after that operation commits, so termination can
    affect only the current claim. The caller records an interrupted inspection
    as a review case using its existing error path.
    """

    def __init__(self, root, maximum_objects=64, maximum_private_bytes=384 * 1024**2):
        self.root = Path(root)
        self.maximum_objects = maximum_objects
        self.maximum_private_bytes = maximum_private_bytes
        self.process = None
        self.log = None
        self.reader = None
        self.responses = None
        self.completed = 0

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.close()

    def start(self):
        """Start a fresh protocol reader and a parser with its own diagnostic log."""
        self.log_path = self.root / 'logs/parser-sessions' / (uuid.uuid4().hex + '.log')
        self.log_path.parent.mkdir(parents=True, exist_ok=True)
        self.log = self.log_path.open('ab', buffering=0)
        command = [sys.executable, '-u', __file__, '--root', str(self.root)]
        try:
            self.process = subprocess.Popen(command, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                                            stderr=self.log, creationflags=0x08000000 if sys.platform == 'win32' else 0)
        except BaseException:
            self.log.close()
            raise
        self.responses = Queue()
        self.completed = 0
        self.reader = Thread(target=self.read_responses, daemon=True)
        self.reader.start()

    def read_responses(self):
        """Drain diagnostics and deliver only framed inspection responses."""
        try:
            for line in self.process.stdout:
                if line.startswith(RESULT_PREFIX):
                    try:
                        self.responses.put(json.loads(line[len(RESULT_PREFIX):]))
                    except ValueError:
                        self.responses.put(None)
                else:
                    self.log.write(line)
        finally:
            self.responses.put(None)

    def inspect(self, sha, timeout=300):
        """Return the existing inspector exit code or fail the current claim.

        The deadline includes process startup and imports. A dead or stalled
        parser is terminated and cannot deliver a response to the next request.
        """
        import psutil
        if self.process is None:
            self.start()
        try:
            self.process.stdin.write((sha + '\n').encode('ascii'))
            self.process.stdin.flush()
            response = self.responses.get(timeout=timeout)
            if response is None or response.get('sha') != sha:
                raise RuntimeError('Parser exited without the expected response; see ' + str(self.log_path))
            code = response['returncode']
        except Empty:
            self.close(force=True)
            raise TimeoutError('Object parser exceeded its deadline; see ' + str(self.log_path)) from None
        except OSError as error:
            self.close(force=True)
            raise RuntimeError('Parser pipe failed; see ' + str(self.log_path)) from error
        except BaseException:
            self.close(force=True)
            raise
        self.completed += 1
        try:
            private = psutil.Process(self.process.pid).memory_info().private
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            private = self.maximum_private_bytes
        if self.completed >= self.maximum_objects or private >= self.maximum_private_bytes:
            self.close()
        return code

    def close(self, force=False):
        """Stop the owned parser, drain its reader and release handles before reuse."""
        if self.process is None:
            return
        if self.process.poll() is None:
            if force:
                self.process.terminate()
            else:
                try:
                    self.process.stdin.close()
                except OSError:
                    if self.process.poll() is None:
                        raise
            try:
                self.process.wait(timeout=10)
            except subprocess.TimeoutExpired:
                self.process.kill()
                self.process.wait()
        self.reader.join(timeout=10)
        if self.reader.is_alive():
            raise RuntimeError('Parser protocol reader did not finish')
        self.process.stdout.close()
        if not self.process.stdin.closed:
            try:
                self.process.stdin.close()
            except OSError:
                if self.process.poll() is None:
                    raise
        self.log.close()
        self.process = None


def serve(root):
    """Inspect requests serially and acknowledge only completed operations."""
    from bootstrap import setup
    setup(root)
    sys.path.insert(0, str(Path(__file__).parent))
    from workers import inspect_one
    for line in sys.stdin:
        sha = line.strip()
        if len(sha) != 64 or any(character not in '0123456789abcdef' for character in sha):
            raise ValueError('Invalid inspection request')
        print('Inspecting ' + sha, file=sys.stderr, flush=True)
        with redirect_stdout(sys.stderr):
            code = inspect_one(root, sha)
        response = json.dumps({'sha': sha, 'returncode': code}).encode('ascii')
        sys.stdout.buffer.write(RESULT_PREFIX + response + b'\n')
        sys.stdout.buffer.flush()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root', type=Path, required=True)
    serve(parser.parse_args().root)
