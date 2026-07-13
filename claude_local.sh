#!/bin/bash
#ANTHROPIC_AUTH_TOKEN=ollama ANTHROPIC_BASE_URL=http://10.200.117.19:11434 ANTHROPIC_API_KEY="" claude --model qwen3.5:latest --channels plugin:telegram@claude-plugins-official
#ANTHROPIC_AUTH_TOKEN=ollama
ANTHROPIC_API_KEY="" \
	ANTHROPIC_BASE_URL=http://10.200.117.19:4000 \
	claude --model ollama/glm-4.7-flash-32k:latest --channels plugin:telegram@claude-plugins-official --verbose
