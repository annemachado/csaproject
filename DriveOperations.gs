//Pesquisa ou cria uma Folder no Drive
function getOrCreateFolder(folderName, parentFolder) {
  var folders = (parentFolder ? parentFolder : DriveApp).getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return (parentFolder ? parentFolder : DriveApp).createFolder(folderName);
  }
}

//Cria um arquivo no Drive, a partir de um blob, na pasta parentFolder
function createFileInDrive(blob, parentFolder) {
  var file = parentFolder.createFile(blob);
  return file; // Retorna o arquivo
}
