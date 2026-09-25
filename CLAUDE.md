# Arbejdsregler for Claude Code i dette repo

Disse regler gælder i alle sessioner: terminal, desktop-app og cloud (claude.ai/code).

## Sprog

- Svar, commit-beskeder og kommentarer skrives på dansk.

## Commit og push

- Når en opgave er færdig og testene er grønne: commit med en beskrivende
  dansk commit-besked og push med `git push -u origin <branch>`.
- Arbejd altid på en feature-branch. Lav ikke commits direkte på `master`,
  medmindre brugeren beder om det.
- Del store bunker af ændringer op i flere commits, så historikken giver mening.

## Beskyt committet arbejde

Gamle, ikke-committede ændringer må aldrig overskrive nyere, committede ændringer.

Før du committer lokale ændringer, som kan være ældre end det, der ligger på GitHub:

1. Kør `git fetch origin`.
2. Find filer, der er ændret lokalt OG har nyere commits på `origin`.
   Rør ikke ved dem endnu.
3. Commit og push kun de filer, der ikke er i konflikt.
4. For filer i konflikt: vis brugeren forskellen med `git diff`, og lad
   brugeren vælge, hvilken version der skal beholdes.

## Kommandoer, der kræver udtrykkelig tilladelse fra brugeren

Kør aldrig disse uden at brugeren har sagt ja i samme samtale:

- `git push --force` og `git push --force-with-lease`
- `git reset --hard`
- `git checkout -- <fil>` og `git restore <fil>` (sletter ikke-committede ændringer)
- `git clean`
- `git rebase` på en branch, der allerede er pushet
- `git stash drop` og `git stash clear`

Hvis en opgave kræver en af dem: forklar hvorfor, og vent på svar.

## Projektstruktur

- `src/` – Python-værktøjerne (extractor, krydstjek, intern_tjek, semantik, GUI).
- `docs/` – dokumentation.
- `testdata/` – Word-testdokumenter. Filer med `Test `-præfiks er testfiler.
- `arkiv/` – gamle versioner, rør ikke ved dem.
- `Tool til udbudstidsplaner/` – selvstændigt tidsplan-værktøj.
- `.env` og andre filer i `.gitignore` må aldrig committes.
