# Release Checklist

## 1. Version Bump — update version in these files:
1. `autocore/__init__.py`
2. `setup.py`
3. `docs/source/conf.py`
4. `README.md`
5. `docs/source/introduction.rst`
6. `changelog.rst`

## 2. Release Steps:
1. Commit all changes.
2. Push to GitHub.
3. Create new release on GitHub with new tag.
4. Verify GitHub Actions published to PyPI successfully.
5. Verify ReadTheDocs updated automatically via webhook. If not, trigger manual build from ReadTheDocs dashboard.