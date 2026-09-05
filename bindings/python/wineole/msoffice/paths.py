import os
import re
import subprocess

# Z:\..., C:/..., \\server\share\...
WINDOWS_SHAPED = re.compile(r'^(?:[A-Za-z]:[\\/]|\\\\)')


class Paths:
    """Linux <-> Wine path conversion, and the question of whether
    converting is meaningful at all.

    It is meaningful only when the client and the bridge are on one machine
    looking at one Wine prefix. Convert when that is not true and you get a
    path that silently refers to some other machine's filesystem.
    """

    @staticmethod
    def convertible(client, windows=(os.name == 'nt')):
        """Deliberately keyed off the same loopback test the bridge uses to
        decide whether a token is required. If the two definitions of
        "local" drifted apart, a connection could be remote for
        authentication and local for paths at once.

        A host's own NIC address counts as remote. Enumerating local
        interfaces to notice otherwise would be more code, more edge cases
        (containers, NAT, temporary IPv6 addresses), and would reintroduce
        exactly that split.
        """
        if windows:  # already Windows paths, and no winepath here
            return False
        return client.loopback

    @staticmethod
    def to_wine(path):
        """Linux path -> Wine path. Returns the argument unchanged when it
        already looks like a Windows path, when it is empty, and when
        winepath is unavailable or fails.

        Failing to convert is not fatal -- the caller can write a Windows
        path themselves, and will see that they need to. Raising here would
        turn a recoverable inconvenience into a stopped script.
        """
        if not path:
            return path
        if WINDOWS_SHAPED.match(str(path)):
            return path
        return Paths._run_winepath('-w', path) or path

    @staticmethod
    def to_local(path):
        """Wine path -> Linux path. Same failure stance."""
        if not path:
            return path
        if not WINDOWS_SHAPED.match(str(path)):
            return path
        return Paths._run_winepath('-u', path) or path

    @staticmethod
    def _run_winepath(flag, path):
        try:
            proc = subprocess.run(
                ['winepath', flag, str(path)],
                capture_output=True, text=True, check=False,
            )
        except OSError:
            # FileNotFoundError (no winepath on PATH) is an OSError, and so
            # is every other way spawning can fail here.
            return None
        if proc.returncode != 0:
            return None
        out = proc.stdout.strip()
        return out or None
