import React, { useState, useEffect } from 'react';
import { 
  Container, 
  Paper, 
  Typography, 
  Button, 
  CircularProgress,
  Box,
  Card,
  CardContent,
  Divider,
  List,
  ListItem,
  ListItemText,
  ListItemIcon,
  Stack,
  useTheme,
  FormControl,
  InputLabel,
  Select,
  MenuItem,
  Grid,
  Tabs,
  Tab
} from '@mui/material';
import CheckCircleOutlineIcon from '@mui/icons-material/CheckCircleOutline';
import axios from 'axios';
import SprintAnalysis from './components/SprintAnalysis';

function App() {
  const theme = useTheme();
  const [loading, setLoading] = useState(false);
  const [sprintData, setSprintData] = useState(null);
  const [error, setError] = useState(null);
  const [boards, setBoards] = useState([]);
  const [sprints, setSprints] = useState([]);
  const [selectedBoard, setSelectedBoard] = useState('');
  const [selectedSprint, setSelectedSprint] = useState('');
  const [activeTab, setActiveTab] = useState(0);

  useEffect(() => {
    fetchBoards();
  }, []);

  useEffect(() => {
    if (selectedBoard) {
      fetchSprints(selectedBoard);
    } else {
      setSprints([]);
      setSelectedSprint('');
    }
  }, [selectedBoard]);

  const fetchBoards = async () => {
    try {
      const response = await axios.get('http://localhost:5000/api/boards');
      setBoards(response.data);
      if (response.data.length > 0) {
        setSelectedBoard(response.data[0].id);
      }
    } catch (err) {
      setError(err.response?.data?.error || 'Failed to fetch boards');
    }
  };

  const fetchSprints = async (boardId) => {
    try {
      const response = await axios.get(`http://localhost:5000/api/sprints?boardId=${boardId}`);
      setSprints(response.data);
      if (response.data.length > 0) {
        setSelectedSprint(response.data[0].id);
      }
    } catch (err) {
      setError(err.response?.data?.error || 'Failed to fetch sprints');
    }
  };

  const fetchSprintReport = async () => {
    if (!selectedBoard || !selectedSprint) {
      setError('Please select both a board and a sprint');
      return;
    }

    setLoading(true);
    setError(null);
    try {
      const response = await axios.get(`http://localhost:5000/api/sprint-report?boardId=${selectedBoard}&sprintId=${selectedSprint}`);
      setSprintData(response.data);
    } catch (err) {
      setError(err.response?.data?.error || 'Failed to fetch sprint report');
    } finally {
      setLoading(false);
    }
  };

  const renderAchievements = (achievements, storyAssignments) => {
    if (!achievements) return null;

    const assignmentSections = storyAssignments.split('\n\n').filter(section => section.trim());
    const storyMappings = {};
    
    assignmentSections.forEach(section => {
      const [title, ...stories] = section.split('\n');
      storyMappings[title] = stories.map(story => {
        const match = story.trim().match(/^- (.*?): (.*)$/);
        return match ? { key: match[1], summary: match[2] } : null;
      }).filter(Boolean);
    });

    const sections = achievements.split('\n\n').filter(section => section.trim());
    
    return (
      <Stack spacing={3}>
        {sections.map((section, index) => {
          const [title, firstAchievement, ...remainingAchievements] = section.split('\n');
          const assignedStories = storyMappings[title] || [];
          
          // Filter out empty achievements
          const filteredAchievements = remainingAchievements.filter(achievement => 
            achievement.trim() && achievement.trim() !== '-'
          );
          
          return (
            <Box 
              key={index}
              sx={{
                p: 3,
                borderRadius: 2,
                backgroundColor: 'rgba(0, 0, 0, 0.02)',
                transition: 'all 0.2s ease-in-out',
                '&:hover': {
                  backgroundColor: 'rgba(0, 0, 0, 0.04)',
                  transform: 'translateY(-2px)',
                  boxShadow: '0 4px 20px rgba(0,0,0,0.1)'
                }
              }}
            >
              <Typography 
                variant="h6" 
                gutterBottom 
                sx={{ 
                  color: 'primary.main',
                  fontWeight: 'bold',
                  background: `linear-gradient(45deg, ${theme.palette.primary.main}, ${theme.palette.primary.light})`,
                  WebkitBackgroundClip: 'text',
                  WebkitTextFillColor: 'transparent',
                  mb: 2
                }}
              >
                {title}
              </Typography>
              
              <Typography 
                variant="subtitle1" 
                sx={{ 
                  color: 'text.secondary',
                  fontStyle: 'italic',
                  mb: 3,
                  pl: 2,
                  borderLeft: `3px solid ${theme.palette.primary.main}`,
                  py: 1
                }}
              >
                {firstAchievement?.trim().replace('- ', '')}
              </Typography>

              <List sx={{ py: 0 }}>
                {filteredAchievements.map((achievement, achievementIndex) => {
                  const achievementText = achievement.trim().replace('- ', '');
                  return (
                    <ListItem 
                      key={achievementIndex} 
                      sx={{ 
                        py: 1,
                        pl: 2,
                        transition: 'all 0.2s ease-in-out',
                        '&:hover': {
                          backgroundColor: 'rgba(0, 0, 0, 0.02)',
                          transform: 'translateX(8px)'
                        }
                      }}
                    >
                      <ListItemIcon sx={{ minWidth: 32 }}>
                        <CheckCircleOutlineIcon 
                          color="primary" 
                          sx={{ 
                            fontSize: '1.2rem',
                            filter: 'drop-shadow(0 2px 4px rgba(0,0,0,0.1))'
                          }} 
                        />
                      </ListItemIcon>
                      <ListItemText 
                        primary={achievementText}
                        primaryTypographyProps={{
                          variant: 'body2',
                          color: 'text.primary',
                          sx: { lineHeight: 1.6 }
                        }}
                      />
                    </ListItem>
                  );
                })}
              </List>

              {assignedStories.length > 0 && (
                <Box sx={{ mt: 3, pl: 2 }}>
                  <Typography 
                    variant="subtitle2" 
                    sx={{ 
                      color: 'text.secondary',
                      fontWeight: 'medium',
                      mb: 2,
                      textTransform: 'uppercase',
                      letterSpacing: '1px'
                    }}
                  >
                    Assigned Stories
                  </Typography>
                  <List sx={{ py: 0 }}>
                    {assignedStories.map((story, storyIndex) => (
                      <ListItem 
                        key={storyIndex} 
                        sx={{ 
                          py: 1,
                          pl: 2,
                          mb: 1,
                          borderRadius: 2,
                          transition: 'all 0.2s ease-in-out',
                          '&:hover': {
                            backgroundColor: 'rgba(0, 0, 0, 0.02)',
                            transform: 'translateX(8px)'
                          }
                        }}
                      >
                        <ListItemText 
                          primary={
                            <Box sx={{ display: 'flex', flexDirection: 'column', gap: 0.5 }}>
                              <Typography
                                sx={{ 
                                  fontFamily: 'monospace',
                                  fontSize: '0.8rem',
                                  color: 'primary.main',
                                  letterSpacing: '0.5px'
                                }}
                              >
                                {story.key}
                              </Typography>
                              <Typography
                                variant="body2"
                                color="text.secondary"
                                sx={{ lineHeight: 1.4 }}
                              >
                                {story.summary}
                              </Typography>
                            </Box>
                          }
                        />
                      </ListItem>
                    ))}
                  </List>
                </Box>
              )}
            </Box>
          );
        })}
      </Stack>
    );
  };

  const handleTabChange = (event, newValue) => {
    setActiveTab(newValue);
  };

  return (
    <Container maxWidth="md" sx={{ py: 6 }}>
      <Paper 
        elevation={3} 
        sx={{ 
          p: 4, 
          mb: 4, 
          borderRadius: 3,
          background: `linear-gradient(135deg, ${theme.palette.background.paper} 0%, ${theme.palette.background.default} 100%)`,
          boxShadow: '0 8px 32px rgba(0,0,0,0.1)'
        }}
      >
        <Typography 
          variant="h4" 
          component="h1" 
          gutterBottom 
          sx={{ 
            color: 'primary.main',
            fontWeight: 'bold',
            background: `linear-gradient(45deg, ${theme.palette.primary.main}, ${theme.palette.primary.light})`,
            WebkitBackgroundClip: 'text',
            WebkitTextFillColor: 'transparent',
            mb: 2
          }}
        >
          Sprint Report Generator
        </Typography>

        <Tabs 
          value={activeTab} 
          onChange={handleTabChange} 
          sx={{ mb: 3 }}
          indicatorColor="primary"
          textColor="primary"
        >
          <Tab label="Sprint Report" />
          <Tab label="Sprint Analysis" />
        </Tabs>

        {activeTab === 0 ? (
          <>
            <Typography 
              variant="body1" 
              color="text.secondary" 
              paragraph
              sx={{ 
                fontSize: '1.1rem',
                lineHeight: 1.6,
                mb: 3
              }}
            >
              Generate a detailed report of the last closed sprint, including AI-generated subgoals and story assignments.
            </Typography>

            <Grid container spacing={2} sx={{ mb: 3 }}>
              <Grid item xs={12} md={6}>
                <FormControl 
                  fullWidth 
                  sx={{ 
                    '& .MuiOutlinedInput-root': {
                      borderRadius: 2,
                      backgroundColor: 'rgba(255, 255, 255, 0.9)',
                      '&:hover': {
                        backgroundColor: 'rgba(255, 255, 255, 1)',
                      },
                    }
                  }}
                >
                  <InputLabel id="board-select-label">Select Board</InputLabel>
                  <Select
                    labelId="board-select-label"
                    id="board-select"
                    value={selectedBoard}
                    label="Select Board"
                    onChange={(e) => setSelectedBoard(e.target.value)}
                  >
                    {boards.map((board) => (
                      <MenuItem key={board.id} value={board.id}>
                        {board.name}
                      </MenuItem>
                    ))}
                  </Select>
                </FormControl>
              </Grid>
              <Grid item xs={12} md={6}>
                <FormControl 
                  fullWidth 
                  sx={{ 
                    '& .MuiOutlinedInput-root': {
                      borderRadius: 2,
                      backgroundColor: 'rgba(255, 255, 255, 0.9)',
                      '&:hover': {
                        backgroundColor: 'rgba(255, 255, 255, 1)',
                      },
                    }
                  }}
                >
                  <InputLabel id="sprint-select-label">Select Sprint</InputLabel>
                  <Select
                    labelId="sprint-select-label"
                    id="sprint-select"
                    value={selectedSprint}
                    label="Select Sprint"
                    onChange={(e) => setSelectedSprint(e.target.value)}
                    disabled={!selectedBoard}
                  >
                    {sprints.map((sprint) => (
                      <MenuItem key={sprint.id} value={sprint.id}>
                        {sprint.name} ({new Date(sprint.startDate).toLocaleDateString()} - {new Date(sprint.endDate).toLocaleDateString()})
                      </MenuItem>
                    ))}
                  </Select>
                </FormControl>
              </Grid>
            </Grid>

            <Box sx={{ display: 'flex', gap: 2, mt: 2 }}>
              <Button 
                variant="contained" 
                onClick={fetchSprintReport}
                disabled={loading || !selectedBoard || !selectedSprint}
                sx={{ 
                  borderRadius: 2,
                  px: 4,
                  py: 1.5,
                  fontSize: '1.1rem',
                  textTransform: 'none',
                  boxShadow: '0 4px 14px rgba(0,0,0,0.1)',
                  transition: 'all 0.2s ease-in-out',
                  '&:hover': {
                    transform: 'translateY(-2px)',
                    boxShadow: '0 6px 20px rgba(0,0,0,0.15)'
                  }
                }}
              >
                {loading ? <CircularProgress size={24} /> : 'Generate Report'}
              </Button>

              {sprintData && (
                <Button
                  variant="outlined"
                  onClick={() => {
                    const url = `http://localhost:5000/api/sprint-report/download?boardId=${selectedBoard}&sprintId=${selectedSprint}`;
                    window.open(url, '_blank');
                  }}
                  sx={{
                    borderRadius: 2,
                    px: 4,
                    py: 1.5,
                    fontSize: '1.1rem',
                    textTransform: 'none',
                    borderWidth: 2,
                    '&:hover': {
                      borderWidth: 2,
                      transform: 'translateY(-2px)',
                      boxShadow: '0 6px 20px rgba(0,0,0,0.1)'
                    }
                  }}
                >
                  Download Word Document
                </Button>
              )}
            </Box>
          </>
        ) : (
          <SprintAnalysis />
        )}
      </Paper>

      {error && (
        <Paper 
          elevation={3} 
          sx={{ 
            p: 3, 
            mb: 4, 
            bgcolor: '#ffebee', 
            borderRadius: 2,
            border: '1px solid #ffcdd2'
          }}
        >
          <Typography color="error">{error}</Typography>
        </Paper>
      )}

      {sprintData && activeTab === 0 && (
        <Card 
          elevation={3} 
          sx={{ 
            borderRadius: 3,
            overflow: 'hidden',
            boxShadow: '0 8px 32px rgba(0,0,0,0.1)'
          }}
        >
          <CardContent sx={{ p: 4 }}>
            <Typography 
              variant="h5" 
              gutterBottom 
              sx={{ 
                color: 'primary.main',
                fontWeight: 'bold',
                background: `linear-gradient(45deg, ${theme.palette.primary.main}, ${theme.palette.primary.light})`,
                WebkitBackgroundClip: 'text',
                WebkitTextFillColor: 'transparent',
                mb: 1
              }}
            >
              {sprintData.sprint_name}
            </Typography>
            <Typography 
              variant="subtitle1" 
              color="text.secondary" 
              gutterBottom
              sx={{ mb: 4 }}
            >
              {new Date(sprintData.start_date).toLocaleDateString()} - {new Date(sprintData.end_date).toLocaleDateString()}
            </Typography>
            
            <Box sx={{ my: 4 }}>
              <Typography 
                variant="h6" 
                gutterBottom 
                sx={{ 
                  color: 'primary.main',
                  fontWeight: 'bold',
                  mb: 2
                }}
              >
                Sprint Goal
              </Typography>
              <Typography 
                variant="body1" 
                paragraph 
                sx={{ 
                  color: 'text.primary',
                  fontWeight: 500,
                  fontSize: '1.1rem',
                  lineHeight: 1.6,
                  p: 2,
                  borderRadius: 2,
                  backgroundColor: 'rgba(0, 0, 0, 0.02)'
                }}
              >
                {sprintData.sprint_goal}
              </Typography>
            </Box>

            <Divider sx={{ my: 4 }} />

            {renderAchievements(sprintData.achievements, sprintData.story_assignments)}
          </CardContent>
        </Card>
      )}
    </Container>
  );
}

export default App; 