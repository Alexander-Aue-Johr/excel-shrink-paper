import json
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch

import run_bounded

class ReserveTests(unittest.TestCase):
    def check_reserve(self, ram, commit, expected):
        class Process:
            done = False
            def poll(self):
                return 0 if self.done else None
            def wait(self, timeout):
                self.done = True
        process = Process()
        class Job:
            terminated = 0
            def memory_events(self):
                return []
            def terminate(self):
                self.terminated += 1
                process.done = True
            def counters(self):
                return {}
            def process_ids(self):
                return []
        job = Job()
        args = SimpleNamespace(reserve_mib=8, timeout=10)
        with tempfile.TemporaryDirectory() as directory:
            state = Path(directory) / 'state.json'
            state.write_text(json.dumps({'phase':'opening'}))
            memory = SimpleNamespace(virtual_memory=lambda:SimpleNamespace(available=ram))
            with patch.object(run_bounded,'psutil',memory,create=True), patch.object(run_bounded,'available_commit_bytes',return_value=commit), patch.object(run_bounded,'process_memory',return_value=(0,0)):
                status, _, counters = run_bounded.monitor(job,process,state,args,set(),0)
        self.assertEqual(status,expected)
        self.assertEqual(job.terminated,1)
        self.assertEqual(counters['minimum_free_ram_bytes'],ram)
        self.assertEqual(counters['minimum_commit_headroom_bytes'],commit)

    def test_commit_reserve_stops_owned_job(self):
        self.check_reserve(64*2**30,0,'system_commit_reserve')

    def test_physical_reserve_stops_owned_job(self):
        self.check_reserve(0,64*2**30,'system_ram_reserve')

if __name__=='__main__':
    unittest.main()
