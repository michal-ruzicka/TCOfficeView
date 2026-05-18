# TODO

- [x] Update README.md to be better useable for users and developers.
- [x] Add info page if MS Office is not available on the target system.
- [x] Add some open source permissive license.
- [x] Create standard Total Commander ZIP distribution archive.
- [x] Add configuration via ini file.
  - [x] System wide ini file in Total Commander installation directory.
    - [x] User specific ini file in `%AppData%\GHISLER` TC config directory that has precedence over the global config.
  - [x] Make here logging optional (to not overfill our expensive disks) with configuration of log path.
- [x] Create private repository on GitHub.
  - [ ] Setup standard development process – merge requests from `develop` branch to protected `main` branch.
  - [ ] Setup standard build and release process with GitHub actions on tags on `main` branch.
  - [ ] Make the repo public.
- [ ] Add full embedded mode instead of OLE preview (to allow in Word paged layout instead of web view mode).
  - [ ] Add configuration per file type (Word/Excel/PowerPoint) to use quick of full mode.