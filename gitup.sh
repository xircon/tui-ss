#!/usr/bin/env bash

cd ~/scripts/tui-ss
git status -sb
git add -A
git commit -m "$1"
GIT_SSH_COMMAND='ssh -F /dev/null -o IdentitiesOnly=yes -i /home/xircon/.ssh/id_rsa -o StrictHostKeyChecking=accept-new' git push -u origin main
echo $(date)
