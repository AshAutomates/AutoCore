# Release Checklist

## Version Bump — update version in these files:
- [ ] `autocore/__init__.py`
- [ ] `setup.py`
- [ ] `docs/source/conf.py`
- [ ] `README.md`
- [ ] `docs/source/introduction.rst`
- [ ] `changelog.rst`

## Release Steps:
- [ ] Commit all changes.
- [ ] Push to GitHub.
- [ ] Create new release on GitHub with new tag.
- [ ] Verify GitHub Actions published to PyPI successfully.
- [ ] Verify ReadTheDocs updated automatically via webhook. If not, trigger manual build from ReadTheDocs dashboard.