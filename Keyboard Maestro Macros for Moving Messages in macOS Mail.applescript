-- Based on @MacSparky['s idea](https://www.macsparky.com/blog/2021/8/my-free-apple-mail-seminar) of using the `⌃M` shortkey to trigger a Conflict Pallete on Keyboard Maestro that will filter Macros to Move Messages to specific folders on macOS Mail.app

use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

(*
How this script works:
 0. Please set the variables bellow
 1. Gets the list of Accounts in Mail.App
 2. Loops for each account
 	2.1. Creates a list of every mailbox on that account
 	2.2. Asks for user selection of desired mailboxes on that list
 	2.3. Creates the corresponding set of Macros for Keyboard Maestro
*)

-- Parameters to set up
## The Keyboard Maestro Macro Group Name (will be created if there is no such group yet).
set theMacroGroupName to "Mail Moves"

## The Keyboard Maestro Trigger
-- Keyboard Maestro sets the contents of its items properties via xml dictionaries. It's easier to copy and paste its contents from a example in Keyboard Maestro Editor window and its menu (Edit > Copy as > Copy as XML)
-- Here's the XML for using ⌃M as a trigger for this macro
set theXMLfortheTrigger to "<dict>
											<key>FireType</key>
											<string>Pressed</string>
											<key>KeyCode</key>
											<integer>46</integer>
											<key>MacroTriggerType</key>
											<string>HotKey</string>
											<key>Modifiers</key>
											<integer>4096</integer>
										</dict>"



# The script itself

-- Scafolding, do not bother
set theList to {}

tell application "Mail"
	
	-- Creates account List
	set accountList to every account
	
	-- Loops for each account
	repeat with eachAccount in accountList
		set accountName to name of eachAccount
		
		-- Fetches the names for account mailboxes (folders) to create a list
		set accountMailboxes to every mailbox of eachAccount
		repeat with eachMailbox in accountMailboxes
			set the end of theList to (name of eachMailbox)
		end repeat
		
		-- Chooses desired Mailboxes from that list		
		set selectedMailboxes to choose from list theList with title ("Pick the desired mailboxes from " & accountName) with multiple selections allowed
		
		-- This if statement to allow for no selection to be passed via the Cancel button
		if (count of selectedMailboxes) > 0 then
			
			-- Loops the selected Mailboxes to create their corresponding Macros on Keyboard Maestro
			repeat with theMailbox in selectedMailboxes
				
				set theMailbox to theMailbox
				set theAccount to (name of eachAccount)
				
				set theMacroName to "Move to " & theAccount & " - " & theMailbox
				
				-- Will use the Keyboard Maestro Editor NOT the Engine
				tell application "Keyboard Maestro"
					
					-- Creates a Macro Group if there isn`t one named as intended.
					if macro group theMacroGroupName exists then
						set theMacroGroup to macro group theMacroGroupName
					else
						
						-- Will create a New Macro Group to be used only within Mail.app
						set theMacroGroup to make new macro group with properties {name:theMacroGroupName, available application xml:"<dict>
					<key>Targeting</key>
					<string>Included</string>
					<key>TargetingApps</key>
					<array>
						<dict>
							<key>BundleIdentifier</key>
							<string>com.apple.mail</string>
							<key>Name</key>
							<string>Mail</string>
							<key>NewFile</key>
							<string>/System/Applications/Mail.app</string>
						</dict>
					</array>
				</dict>"}
					end if
					
					-- Creates the Macro	
					
					tell theMacroGroup
						
						-- Creates new macro
						set theMacro to make new macro with properties {name:theMacroName}
						
						-- Uses a tell command to address the new Macro
						tell theMacro
							
							-- Creates the Trigger from the XML provided above
							make new trigger with properties {xml:theXMLfortheTrigger}
							
							-- Creates the AppleScript Action
							make new action with properties {xml:"<dict>
				<key>DisplayKind</key>
				<string>None</string>
				<key>HonourFailureSettings</key>
				<true/>
				<key>IncludeStdErr</key>
				<false/>
				<key>MacroActionType</key>
				<string>ExecuteAppleScript</string>
				<key>Path</key>
				<string></string>
				<key>Text</key>
				<string>tell application \"Mail\"
		    set selMessage to selection
		    set thisMessage to item 1 of selMessage
		    set mailbox of thisMessage to mailbox \"" & theMailbox & "\" of account \"" & theAccount & "\"
		    display notification with title \"Message Moved to " & theAccount & " " & theMailbox & "\"
		end tell</string>
				<key>TimeOutAbortsMacro</key>
				<true/>
				<key>TrimResults</key>
				<true/>
				<key>TrimResultsNew</key>
				<true/>
				<key>UseText</key>
				<true/>
			</dict>"}
						end tell
					end tell
				end tell
			end repeat
			
			-- Resets account list for next Account iteration
			set theList to {}
		else
			
		end if
	end repeat
end tell
