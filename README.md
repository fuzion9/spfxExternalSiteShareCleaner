## external-site-share-cleaner-webpart

Sharepoint Web Part providing the (as of now, missing) ability to remove elements on a sharepoint modern page.  This is useful for removing items such as Office 365 Launcher Bar, Share/Email Buttons, Footers, and any other elements on a page.

This webpart also provides the ability to add custom querySelector searches to find and hide any css/div/buttons/etc.. and have them hidden.

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Build options

gulp clean - TODO
gulp test - TODO
gulp serve - TODO
gulp bundle - TODO
gulp package-solution - TODO
