# Git Workflow Rules

- **Never push directly to `main`**: Never commit or push changes directly to `main`.
- **Feature Branch Workflow**:
  1. Always create and work on a dedicated branch branched from `main`:
     ```bash
     git checkout -b feat/<description>    # for features / optimizations
     git checkout -b fix/<description>     # for bug fixes / CI fixes
     git checkout -b chore/<description>   # for maintenance, tooling, linting
     ```
  2. Implement changes on the feature branch.
  3. **Mandatory Pre-Push Quality Gates** (must pass 100% before committing / pushing):
     - **Python Linting & Formatting** with `ruff`:
       ```bash
       docker run --rm -v $(pwd):/workspace -w /workspace pyxl-test:latest bash -c "/opt/venv/bin/ruff check python/ tests/ benchmarks/ && /opt/venv/bin/ruff format --check python/ tests/ benchmarks/"
       ```
     - **Rust Linting & Formatting** with `cargo fmt` and `cargo clippy`:
       ```bash
       docker run --rm -v $(pwd):/workspace -w /workspace pyxl-test:latest bash -c "cargo fmt --all -- --check && cargo clippy --workspace --all-targets -- -D warnings"
       ```
     - **Test Suite Verification**:
       ```bash
       docker run --rm -v $(pwd):/workspace -w /workspace pyxl-test:latest bash -c "pytest tests/unit/ tests/compatibility/"
       ```
  4. Commit changes to the feature branch.
  5. Push the feature branch to remote `origin`:
     ```bash
     git push -u origin <branch-name>
     ```
  6. Switch to `main`, merge the feature branch, and push `main` to `origin`:
     ```bash
     git checkout main
     git merge <branch-name>
     git push origin main
     ```


