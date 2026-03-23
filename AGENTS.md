# Agent Instructions — dan-m365-graph-mcp

## Testing requirements

### New features
Every new tool or handler function must have accompanying tests before the code is committed.

### Bug fixes
Every bug fix must include a regression test that:
- Fails before the fix is applied
- Passes after the fix
- Is named to clearly describe the bug (e.g. `Regression: AQMk ID must be URL-encoded`)

### Where to put tests
- Tests live alongside source files: `src/tools/foo.test.ts` for `src/tools/foo.ts`
- Prefer pure function tests (extract logic into exported helpers) over integration tests requiring auth/network
- Use Node's built-in `node:test` runner — no extra test framework needed

### Running tests
```
npm test
```

Always run `npm test` and confirm all tests pass before creating a commit.

## Code conventions
- All Graph API paths for Todo must use `encodeURIComponent` on list/task IDs (Exchange-format IDs contain `==` which breaks raw URL paths)
- Never include `$select` in `list_tasks` requests — it causes HTTP 400 from the Graph Todo API for Exchange-format list IDs
- Default timezone is `Pacific/Auckland`
- Auth uses MSAL device code flow; token cached at `TOKEN_CACHE_PATH`
- Env vars: `AZURE_CLIENT_ID`, `AZURE_TENANT_ID`, `TOKEN_CACHE_PATH`
