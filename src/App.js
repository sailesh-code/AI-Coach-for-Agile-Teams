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
  List,
  ListItem,
  ListItemText,
  ListItemIcon,
  Grid,
  FormControl,
  InputLabel,
  Select,
  MenuItem,
  Accordion,
  AccordionSummary,
  AccordionDetails,
  Stepper,
  Step,
  StepLabel,
  Alert
} from '@mui/material';
import ExpandMoreIcon from '@mui/icons-material/ExpandMore';
import WarningIcon from '@mui/icons-material/Warning';
import TrendingUpIcon from '@mui/icons-material/TrendingUp';
import LightbulbIcon from '@mui/icons-material/Lightbulb';
import CloudUploadIcon from '@mui/icons-material/CloudUpload';
import axios from 'axios';

function App() {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [boards, setBoards] = useState([]);
  const [sprints, setSprints] = useState([]);
  const [selectedBoard, setSelectedBoard] = useState('');
  const [selectedSprint, setSelectedSprint] = useState('');
  const [file, setFile] = useState(null);
  const [activeStep, setActiveStep] = useState(0);
  const [analysisResult, setAnalysisResult] = useState(null);

  const steps = ['Select Board & Sprint', 'Upload Excel File', 'Generate Report'];

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

  const handleFileChange = (event) => {
    const selectedFile = event.target.files[0];
    if (selectedFile && selectedFile.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
      setFile(selectedFile);
      setError(null);
      setActiveStep(1);
    } else {
      setError('Please select a valid Excel file (.xlsx)');
      setFile(null);
    }
  };

  const handleGenerateReport = async () => {
    if (!file) {
      setError('Please select a file first');
      return;
    }

    if (!selectedBoard || !selectedSprint) {
      setError('Please select both a board and a sprint');
      return;
    }

    setLoading(true);
    setError(null);
    setActiveStep(2);

    const formData = new FormData();
    formData.append('file', file);
    formData.append('boardId', selectedBoard);
    formData.append('sprintId', selectedSprint);

    try {
      const response = await axios.post('http://localhost:5000/api/sprint-combined-report', formData, {
        headers: {
          'Content-Type': 'multipart/form-data',
        },
        responseType: 'blob'
      });

      // Create a download link
      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', `sprint_report_and_analysis_${selectedSprint}.docx`);
      document.body.appendChild(link);
      link.click();
      link.remove();
      
      // Reset the form
      setFile(null);
      setActiveStep(0);
    } catch (err) {
      setError(err.response?.data?.error || 'Failed to generate report');
      setActiveStep(1);
    } finally {
      setLoading(false);
    }
  };

  const renderSpillOverAnalysis = (data) => (
    <Box>
      <Typography variant="h6" gutterBottom>
        Spilled Stories
      </Typography>
      <List>
        {data.spilled_stories.map((story, index) => (
          <ListItem key={index}>
            <ListItemIcon>
              <WarningIcon color="error" />
            </ListItemIcon>
            <ListItemText
              primary={story.story_id}
              secondary={
                <>
                  <Typography component="span" variant="body2" color="text.primary">
                    Reason: {story.reason}
                  </Typography>
                  <br />
                  <Typography component="span" variant="body2" color="text.secondary">
                    Prevention: {story.prevention_suggestion}
                  </Typography>
                </>
              }
            />
          </ListItem>
        ))}
      </List>
      <Typography variant="h6" gutterBottom sx={{ mt: 2 }}>
        Root Causes
      </Typography>
      <List>
        {data.root_causes.map((cause, index) => (
          <ListItem key={index}>
            <ListItemText primary={cause} />
          </ListItem>
        ))}
      </List>
      <Typography variant="h6" gutterBottom sx={{ mt: 2 }}>
        Recommendations
      </Typography>
      <List>
        {data.recommendations.map((rec, index) => (
          <ListItem key={index}>
            <ListItemText primary={rec} />
          </ListItem>
        ))}
      </List>
    </Box>
  );

  const renderChurnAnalysis = (data) => (
    <Box>
      <Typography variant="h6" gutterBottom>
        High Churn Stories
      </Typography>
      <List>
        {data.high_churn_stories.map((story, index) => (
          <ListItem key={index}>
            <ListItemIcon>
              <TrendingUpIcon color="warning" />
            </ListItemIcon>
            <ListItemText
              primary={story.story_id}
              secondary={
                <>
                  <Typography component="span" variant="body2" color="text.primary">
                    Churn Count: {story.churn_count}
                  </Typography>
                  <br />
                  <Typography component="span" variant="body2" color="text.secondary">
                    Impact: {story.impact}
                  </Typography>
                </>
              }
            />
          </ListItem>
        ))}
      </List>
      <Typography variant="h6" gutterBottom sx={{ mt: 2 }}>
        Velocity Impact
      </Typography>
      <Typography paragraph>{data.velocity_impact}</Typography>
      <Typography variant="h6" gutterBottom>
        Reduction Suggestions
      </Typography>
      <List>
        {data.reduction_suggestions.map((suggestion, index) => (
          <ListItem key={index}>
            <ListItemText primary={suggestion} />
          </ListItem>
        ))}
      </List>
    </Box>
  );

  const renderTeamUtilization = (data) => (
    <Box>
      <Typography variant="h6" gutterBottom>
        Team Utilization
      </Typography>
      <Grid container spacing={2}>
        <Grid item xs={12} md={6}>
          <Card>
            <CardContent>
              <Typography variant="subtitle1" gutterBottom>
                Over-utilized Members
              </Typography>
              <List>
                {data.over_utilized.map((member, index) => (
                  <ListItem key={index}>
                    <ListItemText
                      primary={member.member}
                      secondary={`Utilization: ${member.utilization}%`}
                    />
                  </ListItem>
                ))}
              </List>
            </CardContent>
          </Card>
        </Grid>
        <Grid item xs={12} md={6}>
          <Card>
            <CardContent>
              <Typography variant="subtitle1" gutterBottom>
                Under-utilized Members
              </Typography>
              <List>
                {data.under_utilized.map((member, index) => (
                  <ListItem key={index}>
                      <ListItemText 
                      primary={member.member}
                      secondary={`Utilization: ${member.utilization}%`}
                      />
                    </ListItem>
                ))}
              </List>
            </CardContent>
          </Card>
        </Grid>
      </Grid>
      <Typography variant="h6" gutterBottom sx={{ mt: 2 }}>
        Workload Distribution
      </Typography>
      <Typography paragraph>{data.workload_distribution}</Typography>
      <Typography variant="h6" gutterBottom>
        Optimization Suggestions
      </Typography>
      <List>
        {data.optimization_suggestions.map((suggestion, index) => (
          <ListItem key={index}>
            <ListItemText primary={suggestion} />
          </ListItem>
        ))}
      </List>
    </Box>
  );

  const renderAdditionalImprovements = (data) => (
    <Box>
      <Typography variant="h6" gutterBottom>
        Additional Improvements
      </Typography>
      <List>
        {data.map((improvement, index) => (
          <ListItem key={index}>
            <ListItemIcon>
              <LightbulbIcon color="primary" />
            </ListItemIcon>
            <ListItemText
              primary={improvement.area}
              secondary={
                <>
                  <Typography component="span" variant="body2" color="text.primary">
                    Observation: {improvement.observation}
                  </Typography>
                  <br />
                  <Typography component="span" variant="body2" color="text.secondary">
                    Suggestion: {improvement.suggestion}
                              </Typography>
                </>
                          }
                        />
                      </ListItem>
                    ))}
                  </List>
                </Box>
    );

  return (
    <Container maxWidth="lg" sx={{ py: 4 }}>
      <Paper sx={{ p: 3, mb: 3 }}>
        <Typography variant="h4" gutterBottom align="center">
          Sprint Analysis Report
        </Typography>

        <Box sx={{ mb: 3 }}>
          <Grid container spacing={2}>
              <Grid item xs={12} md={6}>
              <FormControl fullWidth>
                <InputLabel>Board</InputLabel>
                  <Select
                    value={selectedBoard}
                    onChange={(e) => setSelectedBoard(e.target.value)}
                  label="Board"
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
              <FormControl fullWidth>
                <InputLabel>Sprint</InputLabel>
                  <Select
                    value={selectedSprint}
                    onChange={(e) => setSelectedSprint(e.target.value)}
                  label="Sprint"
                  >
                    {sprints.map((sprint) => (
                      <MenuItem key={sprint.id} value={sprint.id}>
                      {sprint.name}
                      </MenuItem>
                    ))}
                  </Select>
                </FormControl>
              </Grid>
            </Grid>
        </Box>

        {error && (
          <Alert severity="error" sx={{ mb: 3 }}>
            {error}
          </Alert>
        )}

        <Box>
          <Stepper activeStep={activeStep} sx={{ mb: 3 }}>
            {steps.map((label) => (
              <Step key={label}>
                <StepLabel>{label}</StepLabel>
              </Step>
            ))}
          </Stepper>

          <Box sx={{ mb: 3 }}>
            <input
              accept=".xlsx"
              style={{ display: 'none' }}
              id="excel-file-upload"
              type="file"
              onChange={handleFileChange}
            />
            <label htmlFor="excel-file-upload">
              <Button 
                variant="contained" 
                component="span"
                startIcon={<CloudUploadIcon />}
                disabled={!selectedBoard || !selectedSprint}
              >
                Upload Excel File
              </Button>
            </label>
            {file && (
              <Typography variant="body2" sx={{ mt: 1 }}>
                Selected file: {file.name}
              </Typography>
            )}
          </Box>

                <Button
            variant="contained"
            onClick={handleGenerateReport}
            disabled={loading || !file || !selectedBoard || !selectedSprint}
            sx={{ mb: 3 }}
          >
            {loading ? <CircularProgress size={24} /> : 'Generate Analysis Report'}
                </Button>

          {analysisResult && (
            <Box>
              <Accordion>
                <AccordionSummary expandIcon={<ExpandMoreIcon />}>
                  <Typography variant="h6">Spill-over Analysis</Typography>
                </AccordionSummary>
                <AccordionDetails>
                  {renderSpillOverAnalysis(analysisResult.spill_over_analysis)}
                </AccordionDetails>
              </Accordion>

              <Accordion>
                <AccordionSummary expandIcon={<ExpandMoreIcon />}>
                  <Typography variant="h6">Churn Analysis</Typography>
                </AccordionSummary>
                <AccordionDetails>
                  {renderChurnAnalysis(analysisResult.churn_analysis)}
                </AccordionDetails>
              </Accordion>

              <Accordion>
                <AccordionSummary expandIcon={<ExpandMoreIcon />}>
                  <Typography variant="h6">Team Utilization</Typography>
                </AccordionSummary>
                <AccordionDetails>
                  {renderTeamUtilization(analysisResult.team_utilization)}
                </AccordionDetails>
              </Accordion>

              <Accordion>
                <AccordionSummary expandIcon={<ExpandMoreIcon />}>
                  <Typography variant="h6">Additional Improvements</Typography>
                </AccordionSummary>
                <AccordionDetails>
                  {renderAdditionalImprovements(analysisResult.additional_improvements)}
                </AccordionDetails>
              </Accordion>
            </Box>
        )}
        </Box>
      </Paper>
    </Container>
  );
}

export default App; 