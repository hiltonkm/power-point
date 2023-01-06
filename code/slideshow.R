# GOAL: Create slideshow from given photos, keep photos aspect ratio

# Libraries --------------------------------------------------------------------
rm(list=ls())

library(tidyverse)
library(officer)
library(png)
library(jpeg)

# Globals ----------------------------------------------------------------------
wd <- 'C:/Users/kathe/OneDrive/Documents/Data Projects/GIT/power-point/'
input <- paste0(wd, 'input/')
output <- paste0(wd, 'output/')
temp <- paste0(wd, 'temp/')
dir.create(temp)

# Load data --------------------------------------------------------------------
# Load empty powerpoint
p <- read_pptx(paste0(input,"slideshow-template.pptx"))

layout_summary(p)
layout_properties(p, layout = "Picture with Caption", master = "Office Theme")

# List of file paths of photos (note: photos not on git due to size)
file_list <- list.files(path=paste0(input,'photos/'), full.names = TRUE)

# Sorting photos by type -------------------------------------------------------
dir.create(paste0(temp, 'photos-png/'))
dir.create(paste0(temp, 'photos-jpg/'))
dir.create(paste0(temp, 'photos-mov/'))
dir.create(paste0(temp, 'photos-heic/'))
# Putting all files into a folder based on type
for (i in 1:length(file_list)) {
  print(i)
  file_path <- file_list[i]
  if (str_detect(file_path, 'jpg')) {
    file.copy(file_path,paste0(temp, 'photos-jpg/',i,'.jpg'))
  } else if (str_detect(file_path, 'png')) {
    file.copy(file_path,paste0(temp, 'photos-png/',i,'.png'))
  } else if (str_detect(file_path, 'heic')) {
    file.copy(file_path,paste0(temp, 'photos-heic/',i,'.heic'))
  } else if (str_detect(file_path, 'mov') | str_detect(file_path, 'MOV')) {
    file.copy(file_path, paste0(temp, 'photos-mov/',i,'.mov'))
  } else {
    print("issue with file!")
    print(file_path)
  }
}

# Converting files into png ----------------------------------------------------
# jpg
for (f in list.files(path=paste0(temp, 'photos-jpg/'), full.names = TRUE)) {
  print(f)
  f_image <- readJPEG(f)
  file_path <- gsub("\\-jpg", "-png", f)
  file_path <- gsub("\\.jpg", ".png", file_path)
  writePNG(f_image,file_path)
}
## Note: For HEIC files, used online converter: https://convertio.co/heic-png/
## will not include .mov files in powerpoint (only want static)

## Note: This takes a little while, might want to find a more efficient way
## to do this process

# Putting photos in slideshow --------------------------------------------------
file_list_png <- list.files(paste0(temp, 'photos-png/'), full.names = TRUE)
ImgZoom <- 6
for (i in 1:length(file_list_png)) {
  print(i)
  t <- readPNG(file_list_png[i],native = TRUE,info = TRUE)
  picture <- external_img(src =file_list_png[i],
                          width = ImgZoom*(attr(t,"info")$dim[1]/attr(t,"info")$dim[2]),
                          height = ImgZoom*(attr(t,"info")$dim[2]/attr(t,"info")$dim[2]))
  p <- p %>%
    add_slide(layout = "Picture with Caption", master = "Office Theme") %>%
    ph_with(value = picture,
            location = ph_location_label("Picture Placeholder 2"),
            use_loc_size=F)
}

# Printing deck
print(p, target = paste0(output, 'final-slideshow.pptx'))

# Removing temp folder
unlink(paste0(wd, 'temp'), recursive = TRUE)
