#!/usr/bin/env node
import { Command } from "commander";
import main from "../src/index.js";

const program = new Command();

program
  .option("-u, --user", "username")
  .option("-p, --password", "password")
  .option("-l, --link", "your login page link")
  .option("-d, --debug", "open debug mode");

program.parse(process.argv);

const options = program.opts();
const debug = options.debug || false;
const [username, password, path] = program.args;

main({ username, password, path, debug });
