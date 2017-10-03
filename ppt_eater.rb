
###########################
### Methods for clarity ###
###########################

def this_directory
  File.dirname(__FILE__)
end

def pacman_file_location
  File.join(this_directory, 'pacman.png').gsub('/', '\\')
end

def presentation_file_location
  File.join(this_directory, 'presentation.pptx')
end

def wait_for_input
  gets
end

def random_direction
  [:west, :east].sample
end

#############################
### Launch PowerPoint     ###
### and open presentation ###
#############################

require 'win32ole'
app = WIN32OLE.new 'PowerPoint.Application'
presentation = app.Presentations.Open(presentation_file_location)

################################
### Wait during presentation ###
################################

wait_for_input
sleep 1.5

########################
### Eat presentation ###
########################

number_of_slides = presentation.Slides.Count
slides = []
number_of_slides.times do |index|
  slides << presentation.Slides(index + 1)
end

slides.reverse.each do |slide|
  slide.select

  shapes = []
  number_of_shapes = slide.Shapes.Count

  if number_of_shapes == 0
    slide.Delete
    next
  end

  number_of_shapes.times do |index|
    shapes << slide.Shapes(index + 1)
  end

  case random_direction
    when :west
      pacman_origin_x = 570
      pacman_origin_y = 25 + Random.rand(400)
      pacman_movement_increment = 600 / number_of_shapes
      pacman = slide.Shapes.AddPicture(pacman_file_location, true, true, pacman_origin_x, pacman_origin_y)
    when :east
      pacman_origin_x = 25
      pacman_origin_y = 25 + Random.rand(400)
      pacman_movement_increment = 0 - (600 / number_of_shapes)
      pacman = slide.Shapes.AddPicture(pacman_file_location, true, true, pacman_origin_x, pacman_origin_y)
      pacman.Flip(0)
  end

  shapes.shuffle.each do |shape|
    # sleep 0.1
    pacman.Left = pacman.Left - pacman_movement_increment
    shape.Delete
  end

  slide.Delete
end

#############################################
### Reconstruct title slide and loop data ###
#############################################

new_slide = presentation.Slides.Add(1, 1)
sleep 0.5
new_slide.Shapes(1).TextFrame.TextRange.Text = 'Thanks for Listening!'
sleep 0.5
loop do
  subtitles = [
      'Joshua Russell', 'Joshua Russell .', 'Joshua Russell ..', 'Joshua Russell ...',
      'russelljoshuaa@gmail.com', 'russelljoshuaa@gmail.com .', 'russelljoshuaa@gmail.com ..', 'russelljoshuaa@gmail.com ...',
      '@russelljoshuaa', '@russelljoshuaa .', '@russelljoshuaa ..', '@russelljoshuaa ...',
      'https://www.linkedin.com/in/joshua-russell-ab139660/', 'https://www.linkedin.com/in/joshua-russell-ab139660/ .', 'https://www.linkedin.com/in/joshua-russell-ab139660/ ..', 'https://www.linkedin.com/in/joshua-russell-ab139660/ ...'
  ]
  subtitles.each do |subtitle|
    new_slide.Shapes(2).TextFrame.TextRange.Text = subtitle
    sleep 0.4
  end
end