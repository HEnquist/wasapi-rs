# Build docs for a new version

- merge changes in `master` into `docs` branch: 
  ```
  git checkout docs
  git pull origin master
  ```

- update docs link in cargo.toml

- clean with `cargo clean`

- build docs: `cargo doc --no-deps`

- copy `target\doc` folder to `docs\x.y.z`

- add, commit and push