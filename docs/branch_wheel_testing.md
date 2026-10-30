# Test a branch wheel without installing Rust locally

The `Test Branch Wheel` workflow (`.github/workflows/test-wheel.yml`) builds the
checked-out source on a Windows x64 GitHub runner. It does not publish to PyPI,
create tags or releases, or change the package version.

## Start a build

- Pushing a commit to `fix/text-fidelity` automatically starts the workflow.
- After the workflow also exists on the repository's default branch, open
  **Actions → Test Branch Wheel → Run workflow** and select the branch to test.
  A new `workflow_dispatch` workflow may not be available in the Actions UI until
  it exists on the default branch; the branch push trigger supports its first run.
- Uncommitted local changes cannot be built by GitHub. Commit and push the desired
  source only when authorized; this document and workflow do not do that for you.

## What the workflow does

1. Builds a release-mode Windows x64 wheel with Python 3.11 and Rust on GitHub.
2. Uploads the wheel and `build-info.json`, recording the source commit, filename,
   build Python, and SHA-256 checksum.
3. Installs that exact wheel in separate Python 3.11 and 3.14 jobs, without building
   the extension again or setting up Rust in the test jobs.
4. Runs the unit, compatibility, and integration Python suites against the installed
   package, including the new remediation regressions.
5. Uploads a pytest log and JUnit report for each test Python, even after test failures.

The wheel artifact is available after the build succeeds even if the regression
suite later fails. Build failures cannot produce a usable wheel. The workflow does
not run the native Rust unit suite or performance benchmarks; those are separate
validation tasks. Passing Python tests is not proof of complete openpyxl compatibility.

## Download and install on Windows

On the workflow run page, download the artifact beginning with
`rpyxl-test-wheel-windows-x64-`. Extract it into a local `wheel-artifact` directory
inside your repository checkout. The wheel and `build-info.json` should be directly
inside that directory. Artifacts expire after 14 days.

From PowerShell in the repository root, create an isolated environment:

```powershell
python -m venv .venv-test
.\.venv-test\Scripts\python.exe -m pip install --upgrade pip
.\.venv-test\Scripts\python.exe -m pip install pytest openpyxl python-calamine psutil tabulate
.\.venv-test\Scripts\python.exe -c "import hashlib, json, pathlib, subprocess, sys; root=pathlib.Path('wheel-artifact'); info=json.loads((root/'build-info.json').read_text()); wheel=root/info['wheel']; assert hashlib.sha256(wheel.read_bytes()).hexdigest()==info['sha256']; subprocess.run([sys.executable,'-m','pip','install','--force-reinstall','--no-deps',str(wheel)],check=True)"
.\.venv-test\Scripts\python.exe -c "import rpyxl, rpyxl._core; print(rpyxl.__version__); print(rpyxl.__file__); print(rpyxl._core.__file__)"
```

No local Rust compiler is needed. Use 64-bit Python: this artifact is not a Windows
ARM64 or 32-bit wheel. Its ABI3 tag provides compatibility from Python 3.9 onward,
but this workflow explicitly exercises only Python 3.11 and 3.14.

Run the prepared regressions from the same checkout revision as the build:

```powershell
.\.venv-test\Scripts\python.exe -m pytest tests/unit tests/compatibility tests/integration -v
```

Check `build-info.json` to distinguish this branch build from the PyPI `0.1.1`
release: the development branch currently has the same version number. Force-install
the downloaded wheel rather than installing `rpyxl` by name from PyPI. Installation
into `.venv-test` leaves your existing environment unchanged.
