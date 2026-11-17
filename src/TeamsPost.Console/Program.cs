// teams-post.exe. Output is Unicode to a console and UTF-8 when redirected (see ConsoleText), so an emoji in a
// message shows as itself.
var output = WinBackground.ConsoleText.Out;
return TeamsPost.Cli.Run(args, new TeamsPost.CliHost { Out = output, Err = WinBackground.ConsoleText.Error });
