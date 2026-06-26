import { useState } from "react";
import { Card, CardContent, CardHeader, CardTitle } from "./ui/card";
import { Button } from "./ui/button";
import { Badge } from "./ui/badge";
import { Checkbox } from "./ui/checkbox";
import { Clock, Calendar, Save, CheckCircle2, X, Globe, AlertCircle } from "lucide-react";
import { toast } from "sonner";

export interface ScheduleConfig {
  schedule: { [day: string]: string[] };
  countries: string[];
}

interface Props {
  onSave: (config: ScheduleConfig) => void;
  initialConfig?: ScheduleConfig;
}

const DAYS_OF_WEEK = [
  { short: "Lun", full: "Lunes" },
  { short: "Mar", full: "Martes" },
  { short: "Mié", full: "Miércoles" },
  { short: "Jue", full: "Jueves" },
  { short: "Vie", full: "Viernes" },
  { short: "Sáb", full: "Sábado" },
  { short: "Dom", full: "Domingo" },
];

const HOURS = Array.from({ length: 24 }, (_, i) => `${i.toString().padStart(2, "0")}:00`);

const COUNTRIES = [
  "Argentina", "Bolivia", "Brasil",
  "Chile", "Colombia", "Ecuador",
  "Paraguay", "Peru", "Uruguay",
];

function CopyToOtherDays({
  sourceDay,
  schedule,
  onApply,
}: {
  sourceDay: string;
  schedule: { [day: string]: string[] };
  onApply: (days: string[]) => void;
}) {
  const [open, setOpen] = useState(false);
  const [selected, setSelected] = useState<string[]>([]);
  const otherDays = DAYS_OF_WEEK.map((d) => d.full).filter((d) => d !== sourceDay);

  const toggle = (day: string) =>
    setSelected((prev) => prev.includes(day) ? prev.filter((d) => d !== day) : [...prev, day]);

  const handleApply = () => {
    if (selected.length === 0) return;
    onApply(selected);
    setSelected([]);
    setOpen(false);
  };

  return (
    <div className="border-t border-white/10 pt-3">
      {!open ? (
        <button
          onClick={() => setOpen(true)}
          className="w-full flex items-center justify-center gap-2 px-4 py-2.5 rounded-lg border border-dashed border-primary/50 bg-primary/10 hover:bg-primary/20 hover:border-primary/80 transition-all duration-150 text-sm font-semibold text-primary cursor-pointer"
        >
          <Calendar className="size-4" />
          Aplicar estos horarios a otros días
        </button>
      ) : (
        <div className="space-y-2.5">
          <p className="text-xs font-semibold text-muted-foreground uppercase tracking-wide">
            Copiar horarios de {sourceDay} a:
          </p>
          <div className="flex flex-wrap gap-2">
            {otherDays.map((day) => {
              const isSelected = selected.includes(day);
              const hasExisting = (schedule[day]?.length || 0) > 0;
              return (
                <button
                  key={day}
                  onClick={() => toggle(day)}
                  className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium border transition-all duration-150
                    ${isSelected
                      ? "bg-primary/40 border-primary text-primary-foreground"
                      : "bg-white/5 border-white/20 hover:bg-white/15 hover:border-white/40 text-foreground"
                    }`}
                >
                  {isSelected && <CheckCircle2 className="size-3" />}
                  {day}
                  {hasExisting && !isSelected && (
                    <span className="text-[10px] text-primary opacity-70">(ya tiene)</span>
                  )}
                </button>
              );
            })}
          </div>
          <div className="flex items-center gap-2">
            <Button size="sm" onClick={handleApply} disabled={selected.length === 0}
              className="h-7 text-xs font-bold uppercase tracking-wide gap-1.5">
              <CheckCircle2 className="size-3.5" />
              Aplicar a {selected.length > 0 ? `${selected.length} día${selected.length !== 1 ? "s" : ""}` : "..."}
            </Button>
            <button onClick={() => { setOpen(false); setSelected([]); }}
              className="text-xs text-muted-foreground hover:text-foreground underline underline-offset-2">
              Cancelar
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

export function AutomationScheduler({ onSave, initialConfig }: Props) {
  const [selectedDay, setSelectedDay] = useState<string | null>(null);
  const [schedule, setSchedule] = useState<{ [day: string]: string[] }>(initialConfig?.schedule ?? {});
  const [selectedCountries, setSelectedCountries] = useState<string[]>(initialConfig?.countries ?? []);

  const toggleHour = (day: string, hour: string) => {
    setSchedule((prev) => {
      const daySchedule = prev[day] || [];
      const exists = daySchedule.includes(hour);
      return { ...prev, [day]: exists ? daySchedule.filter((h) => h !== hour) : [...daySchedule, hour].sort() };
    });
  };

  const toggleCountry = (country: string) =>
    setSelectedCountries((prev) =>
      prev.includes(country) ? prev.filter((c) => c !== country) : [...prev, country]
    );

  const toggleAllCountries = () =>
    setSelectedCountries((prev) => prev.length === COUNTRIES.length ? [] : [...COUNTRIES]);

  const getDayCount = (day: string) => schedule[day]?.length || 0;
  const getTotalSlots = () => Object.values(schedule).reduce((t, h) => t + h.length, 0);

  const clearDaySchedule = (day: string, e: React.MouseEvent | React.KeyboardEvent) => {
    e.stopPropagation();
    setSchedule((prev) => { const next = { ...prev }; delete next[day]; return next; });
  };

  const canSave = getTotalSlots() > 0 && selectedCountries.length > 0;

  const handleSave = () => {
    if (!canSave) {
      toast.error("Configuración incompleta", {
        description: "Seleccioná al menos un horario y un país antes de guardar.",
      });
      return;
    }
    onSave({ schedule, countries: selectedCountries });
    toast.success("Configuración guardada", {
      description: "Ya podés programar el test automático.",
      icon: <CheckCircle2 className="size-5" />,
    });
  };

  return (
    <div className="space-y-5 overflow-y-auto max-h-[75vh] pr-1">
      {/* Days */}
      <Card className="border border-white/20 shadow-xl">
        <CardHeader className="pb-3">
          <div className="flex items-center justify-between">
            <CardTitle className="text-sm uppercase tracking-widest text-muted-foreground font-semibold">
              Días de la semana
            </CardTitle>
            <div className="flex items-center gap-3">
              {getTotalSlots() > 0 && (
                <Badge className="text-xs bg-primary/30 text-primary-foreground border border-primary/40 px-2 py-0.5">
                  {getTotalSlots()} horario{getTotalSlots() !== 1 ? "s" : ""}
                </Badge>
              )}
              {getTotalSlots() > 0 && (
                <button onClick={() => { setSchedule({}); setSelectedDay(null); }}
                  className="text-xs text-primary hover:text-primary/80 underline underline-offset-2 cursor-pointer">
                  Desmarcar todos
                </button>
              )}
            </div>
          </div>
        </CardHeader>
        <CardContent className="space-y-4">
          <div className="grid grid-cols-7 gap-2">
            {DAYS_OF_WEEK.map((day) => {
              const count = getDayCount(day.full);
              const hasHours = count > 0;
              const isOpen = selectedDay === day.full;
              return (
                <div key={day.full} role="button" tabIndex={0}
                  onClick={() => setSelectedDay(isOpen ? null : day.full)}
                  onKeyDown={(e) => e.key === "Enter" && setSelectedDay(isOpen ? null : day.full)}
                  className={`relative flex flex-col items-center gap-1.5 p-3 rounded-xl border transition-all duration-200 cursor-pointer select-none text-center
                    ${isOpen ? "border-primary bg-primary/25 shadow-lg shadow-primary/20"
                      : hasHours ? "border-primary/50 bg-primary/10 hover:bg-primary/20"
                      : "border-white/20 bg-white/5 hover:bg-white/10 hover:border-white/40"}`}
                >
                  <span className="text-[10px] font-bold tracking-widest uppercase text-muted-foreground">{day.short}</span>
                  {hasHours ? (
                    <>
                      <div className="flex flex-col items-center gap-0.5">
                        <div className="flex flex-wrap justify-center gap-[3px] max-w-[40px]">
                          {Array.from({ length: Math.min(count, 6) }).map((_, i) => (
                            <span key={i} className="size-1.5 rounded-full bg-primary/80 inline-block" />
                          ))}
                          {count > 6 && <span className="text-[9px] text-primary font-bold leading-none">+{count - 6}</span>}
                        </div>
                        <span className="text-[9px] text-primary font-semibold">{count} sel.</span>
                      </div>
                      <span role="button" tabIndex={0}
                        onClick={(e) => clearDaySchedule(day.full, e)}
                        onKeyDown={(e) => e.key === "Enter" && clearDaySchedule(day.full, e)}
                        className="text-[10px] text-red-300 hover:text-red-200 hover:underline cursor-pointer leading-none">
                        Limpiar
                      </span>
                    </>
                  ) : (
                    <span className="text-[10px] text-muted-foreground">—</span>
                  )}
                  {isOpen && <span className="absolute -top-1 -right-1 size-2 rounded-full bg-primary" />}
                </div>
              );
            })}
          </div>

          {selectedDay && (
            <div className="rounded-xl border border-white/20 bg-white/5 p-4 space-y-3 animate-in fade-in slide-in-from-top-2 duration-200">
              <div className="flex items-center justify-between">
                <p className="text-sm font-semibold flex items-center gap-2">
                  <Clock className="size-4 text-primary" />
                  Horarios para el {selectedDay}
                </p>
                <div className="flex items-center gap-2">
                  <Button size="sm" variant="secondary" onClick={() => setSelectedDay(null)}
                    className="h-7 gap-1.5 text-xs font-bold uppercase tracking-wide">
                    <CheckCircle2 className="size-3.5" />
                    Listo
                  </Button>
                  <button onClick={() => setSelectedDay(null)}
                    className="rounded-full p-1.5 hover:bg-white/10 transition-colors text-muted-foreground hover:text-foreground"
                    aria-label="Cerrar selector de horas">
                    <X className="size-4" />
                  </button>
                </div>
              </div>
              <div className="grid grid-cols-6 sm:grid-cols-8 md:grid-cols-12 gap-1.5">
                {HOURS.map((hour) => {
                  const picked = schedule[selectedDay]?.includes(hour) || false;
                  return (
                    <button key={hour} onClick={() => toggleHour(selectedDay, hour)}
                      className={`py-1.5 rounded-lg text-xs font-medium transition-all duration-150 border
                        ${picked
                          ? "bg-primary text-primary-foreground border-primary shadow-md shadow-primary/30 scale-105"
                          : "bg-white/8 border-white/20 hover:bg-white/20 hover:border-white/40 text-foreground"}`}>
                      {hour}
                    </button>
                  );
                })}
              </div>
              {schedule[selectedDay]?.length > 0 && (
                <div className="flex flex-wrap gap-1.5 pt-1">
                  {schedule[selectedDay].map((h) => (
                    <Badge key={h} className="text-xs bg-primary/40 text-primary-foreground border border-primary/50">{h}</Badge>
                  ))}
                </div>
              )}
              {schedule[selectedDay]?.length > 0 && (
                <CopyToOtherDays
                  sourceDay={selectedDay}
                  schedule={schedule}
                  onApply={(targetDays) => {
                    setSchedule((prev) => {
                      const next = { ...prev };
                      const hours = prev[selectedDay] || [];
                      targetDays.forEach((d) => {
                        next[d] = Array.from(new Set([...(next[d] || []), ...hours])).sort();
                      });
                      return next;
                    });
                    toast.success(`Horarios copiados a ${targetDays.length} día${targetDays.length !== 1 ? "s" : ""}`, {
                      description: targetDays.join(", "),
                    });
                  }}
                />
              )}
            </div>
          )}
        </CardContent>
      </Card>

      {/* Countries */}
      <Card className="border border-white/20 shadow-xl">
        <CardHeader className="pb-3">
          <div className="flex items-center justify-between">
            <CardTitle className="text-sm uppercase tracking-widest text-muted-foreground font-semibold flex items-center gap-2">
              <Globe className="size-4" />
              Países a testear
            </CardTitle>
            <button onClick={toggleAllCountries}
              className="text-xs text-primary hover:text-primary/80 underline underline-offset-2 cursor-pointer">
              {selectedCountries.length === COUNTRIES.length ? "Desmarcar todos" : "Seleccionar todos"}
            </button>
          </div>
        </CardHeader>
        <CardContent>
          <div className="grid grid-cols-3 gap-3">
            {COUNTRIES.map((country) => {
              const checked = selectedCountries.includes(country);
              return (
                <label key={country}
                  className={`flex items-center gap-2.5 p-3 rounded-xl border cursor-pointer transition-all duration-150
                    ${checked ? "border-primary/50 bg-primary/15 hover:bg-primary/20"
                      : "border-white/20 bg-white/5 hover:bg-white/10 hover:border-white/35"}`}>
                  <Checkbox checked={checked} onCheckedChange={() => toggleCountry(country)}
                    className="border-white/40 data-[state=checked]:bg-primary data-[state=checked]:border-primary" />
                  <span className="text-sm font-medium">{country}</span>
                </label>
              );
            })}
          </div>
          {selectedCountries.length > 0 && (
            <p className="text-xs text-muted-foreground mt-3">
              {selectedCountries.length} de {COUNTRIES.length} países seleccionados
            </p>
          )}
        </CardContent>
      </Card>

      {/* Save */}
      <div className="flex items-center justify-between gap-4 pb-2">
        {!canSave && (
          <div className="flex items-start gap-2 text-xs text-amber-300 bg-amber-500/10 border border-amber-400/20 rounded-lg px-3 py-2.5 flex-1">
            <AlertCircle className="size-3.5 mt-0.5 shrink-0" />
            <span>
              {getTotalSlots() === 0 && selectedCountries.length === 0
                ? "Seleccioná al menos un horario y un país para continuar."
                : getTotalSlots() === 0 ? "Seleccioná al menos un horario para continuar."
                : "Seleccioná al menos un país para continuar."}
            </span>
          </div>
        )}
        <div className="ml-auto">
          <Button onClick={handleSave} disabled={!canSave} size="lg"
            className="gap-2 uppercase tracking-wide text-xs font-bold shadow-lg shadow-primary/30">
            <Save className="size-4" />
            Guardar configuración
          </Button>
        </div>
      </div>
    </div>
  );
}
