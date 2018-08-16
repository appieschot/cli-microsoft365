# spo site classic remove

Remove classic site

## Usage

```sh
spo site classic remove [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --url`| url of the site to remove
`--skipRecycleBin`|set to directly remove the site without moving it to the Recycle Bin
`--confirm`|Don't prompt for confirming removing the file
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online tenant admin site, using the [spo connect](../connect.md) command.

## Remarks


## Examples

Removes a classic site to the Reycle Bin 

```sh
spo site classic remove -u https://tenant.sharepoint.com/sites/demo1
```

Removes a classic site without confirmation prompt 

```sh
spo site classic remove -u https://tenant.sharepoint.com/sites/demo1 --skipRecycleBin --confirm 
```