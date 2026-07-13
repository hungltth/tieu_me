#!/bin/bash
ANTHROPIC_AUTH_TOKEN=ollama ANTHROPIC_BASE_URL=http://10.200.117.19:11434 ANTHROPIC_API_KEY="" claude --model glm-5:cloud --channels plugin:telegram@claude-plugins-official
#ANTHROPIC_AUTH_TOKEN=ollama
#ANTHROPIC_API_KEY="" \
#	ANTHROPIC_BASE_URL=http://10.200.117.19:4000 \
#	claude --model ollama/qwen3.5:latest --channels plugin:telegram@claude-plugins-official
