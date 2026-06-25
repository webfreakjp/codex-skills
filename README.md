# Codex Skills

Personal Codex skill repository.

## Skills

- `pptx-robust-blue-design`: Create or restyle PowerPoint decks in a robust blue guidebook-like visual style.

## Install

Install a skill with `$skill-installer` from this GitHub repository.

```text
$skill-installer install https://github.com/{owner}/codex-skills/tree/main/skills/pptx-robust-blue-design
```

Or specify the repository and path explicitly:

```text
$skill-installer install the skill from {owner}/codex-skills at skills/pptx-robust-blue-design
```

Restart Codex if the newly installed skill does not appear immediately.

## Development

The normal development flow is:

1. Edit the skill in this repository.
2. Symlink the repository copy into the local Codex skills directory.
3. Test locally with Codex and the skill scripts.
4. Push after the skill is working.

Use `$skill-installer` for clean install checks or setup on another machine. You do not need to push and reinstall for every local development change.

### Local symlink setup

Codex discovers user skills from `~/.codex/skills`. During development, point that location at this repository copy:

```bash
rm -rf ~/.codex/skills/pptx-robust-blue-design

ln -s {repo-root}/skills/pptx-robust-blue-design \
  ~/.codex/skills/pptx-robust-blue-design
```

If Codex does not pick up a skill change immediately, start a new thread or restart Codex.
