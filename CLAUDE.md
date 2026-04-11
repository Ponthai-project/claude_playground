# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Purpose

This is a personal sandbox repository (`claude_playground`) used for experimenting with Claude Code. There is no build system, test framework, or package manager.

## Directory structure

- `games/` — browser-based games (HTML files, open directly in a browser)

## Running files

- **HTML files**: Open directly in a browser — no server required. For example, open `games/shooter.html` by double-clicking it or via `start games\shooter.html` on Windows.

## Current contents

### `games/othello.html`
A fully self-contained, single-file browser Othello (Reversi) game written in vanilla JS/HTML/CSS. Key design points:

- **Game logic**: `getFlips`, `validMoves`, `applyMove` — pure functions operating on an 8×8 2D array
- **AI**: Minimax with alpha-beta pruning at depth 4, using a static positional weight matrix (`WEIGHTS`) that heavily values corners
- **Rendering**: Full board re-render on each state change via `render()`; flip animation applied via CSS class injection post-render
- **AI toggle**: White plays as AI when enabled; AI move is triggered via `setTimeout` after the human places a disc
