# Git Workflow Rules

- **Never push directly to `main`**: Never commit or push changes directly to `main`.
- **Feature Branch Workflow**:
  1. Always create and work on a dedicated branch branched from `main`:
     ```bash
     git checkout -b feat/<description>   # for features / optimizations
     git checkout -b fix/<description>    # for bug fixes / CI fixes
     ```
  2. Run all tests and builds on the branch.
  3. Commit and push the branch to remote `origin`.
  4. Merge the branch into `main` (or create a pull request).
